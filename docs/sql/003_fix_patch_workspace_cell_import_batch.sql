-- ============================================================================
-- MIGRATION 003: FIX IMPORT_BATCH_ID PROPAGATION IN WORKBOOK CELL PATCH RPCS
-- ============================================================================
-- Safe, rerunnable migration. Does NOT delete, truncate, or re-import data.
-- Fixes:
--   1. public.patch_workspace_cell: assigns parent workbook_sheets.import_batch_id
--      when inserting a new row and preserves existing workbook_rows.import_batch_id
--      on updates (avoids INSERT ... ON CONFLICT firing BEFORE INSERT triggers
--      with missing/mismatched import_batch_id on existing rows).
--   2. public.patch_workspace_cells_batch: applies the exact same parent
--      import_batch_id propagation, row preservation, and cell-level optimistic
--      concurrency checks atomically across 1..500 edits.
--   3. public.validate_row_import_batch: ensures trigger validation cleanly
--      enforces NEW.import_batch_id IS NOT DISTINCT FROM parent sheet's
--      import_batch_id on both INSERT and UPDATE.
-- ============================================================================

-- 1. Ensure row-level trigger function enforces exact parent import_batch_id match
CREATE OR REPLACE FUNCTION public.validate_row_import_batch()
RETURNS trigger
LANGUAGE plpgsql
SET search_path = pg_catalog, public
AS $$
DECLARE
  v_sheet_import_batch_id uuid;
BEGIN
  SELECT ws.import_batch_id
  INTO v_sheet_import_batch_id
  FROM public.workbook_sheets ws
  WHERE ws.id = NEW.sheet_id;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Parent worksheet % does not exist.', NEW.sheet_id;
  END IF;

  IF NEW.import_batch_id IS DISTINCT FROM v_sheet_import_batch_id THEN
    RAISE EXCEPTION 'INVALID_IMPORT_BATCH: row import batch (%) must match its parent worksheet import batch (%).',
      NEW.import_batch_id,
      v_sheet_import_batch_id;
  END IF;

  RETURN NEW;
END;
$$;

-- 2. Drop any pre-existing overloads cleanly so PostgREST never hits ambiguous signatures
DO $$
DECLARE
  v_sig text;
BEGIN
  FOR v_sig IN
    SELECT p.oid::regprocedure::text
    FROM pg_proc p
    JOIN pg_namespace n ON n.oid = p.pronamespace
    WHERE n.nspname = 'public'
      AND p.proname IN ('patch_workspace_cell', 'patch_workspace_cells_batch')
  LOOP
    EXECUTE 'DROP FUNCTION IF EXISTS ' || v_sig || ';';
  END LOOP;
END;
$$;

-- 3. Recreate patch_workspace_cell with parent import_batch_id inheritance & optimistic concurrency
CREATE OR REPLACE FUNCTION public.patch_workspace_cell(
  p_workspace text,
  p_sheet_id uuid,
  p_row_index integer,
  p_col_index integer,
  p_new_value jsonb,
  p_expected_cell_version bigint
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
DECLARE
  v_user_id uuid;
  v_staff record;
  v_sheet record;
  v_row record;

  v_current_cell_version bigint := 0;
  v_new_cell_version bigint := 1;
  v_new_row_version bigint := 1;
  v_cells jsonb;
  v_cell_versions jsonb;
  v_now timestamptz := now();
BEGIN
  -- 1. Authenticate active staff user & verify workspace permission
  v_user_id := auth.uid();
  IF v_user_id IS NULL THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Authentication required.';
  END IF;

  SELECT user_id, role::text AS role_text, is_active
  INTO v_staff
  FROM public.staff_profiles
  WHERE user_id = v_user_id;

  IF NOT FOUND OR v_staff.is_active IS NOT TRUE THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Inactive or missing staff profile.';
  END IF;

  IF p_workspace IS NULL OR p_workspace NOT IN ('reception', 'doctor') THEN
    RAISE EXCEPTION 'INVALID_WORKSPACE: p_workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % role cannot modify the % workspace.', v_staff.role_text, p_workspace;
  END IF;

  -- 2. Validate coordinates, scalar value, and optimistic concurrency baseline
  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_row_index IS NULL OR p_row_index < 0 OR p_row_index >= 100000 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: row_index must be between 0 and 99999.';
  END IF;

  IF p_col_index IS NULL OR p_col_index < 0 OR p_col_index >= 256 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: col_index must be between 0 and 255.';
  END IF;

  IF p_new_value IS NULL OR jsonb_typeof(p_new_value) NOT IN ('string', 'number') THEN
    RAISE EXCEPTION 'INVALID_CELL_VALUE: Cell value must be a JSON string or number.';
  END IF;

  IF p_expected_cell_version IS NULL OR p_expected_cell_version < 0 THEN
    RAISE EXCEPTION 'CONCURRENCY_REQUIRED: expectedCellVersion must be a nonnegative integer.';
  END IF;

  -- 3. Lock parent worksheet and capture its exact import_batch_id
  SELECT ws.id, ws.workspace, ws.import_batch_id
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet % not found in % workspace.', p_sheet_id, p_workspace;
  END IF;

  -- 4. Lock target row if it already exists
  SELECT wr.id, wr.cells, wr.row_version, wr.cell_versions, wr.import_batch_id
  INTO v_row
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index = p_row_index
  FOR UPDATE;

  IF FOUND THEN
    -- Existing row: enforce cell-level optimistic concurrency and preserve row import_batch_id
    v_current_cell_version := COALESCE((v_row.cell_versions ->> p_col_index::text)::bigint, 0);

    IF v_current_cell_version <> p_expected_cell_version THEN
      RAISE EXCEPTION 'VERSION_CONFLICT: Cell (%, %) expected version %, but current version is %.',
        p_row_index, p_col_index, p_expected_cell_version, v_current_cell_version;
    END IF;

    v_new_cell_version := v_current_cell_version + 1;
    v_new_row_version := COALESCE(v_row.row_version, 0) + 1;

    v_cells := CASE
      WHEN v_row.cells IS NOT NULL AND jsonb_typeof(v_row.cells) = 'array' THEN v_row.cells
      ELSE '[]'::jsonb
    END;

    WHILE jsonb_array_length(v_cells) <= p_col_index LOOP
      v_cells := v_cells || jsonb_build_array('');
    END LOOP;

    v_cells := jsonb_set(v_cells, ARRAY[p_col_index::text], p_new_value, false);

    v_cell_versions := CASE
      WHEN v_row.cell_versions IS NOT NULL AND jsonb_typeof(v_row.cell_versions) = 'object' THEN v_row.cell_versions
      ELSE '{}'::jsonb
    END;

    v_cell_versions := jsonb_set(
      v_cell_versions,
      ARRAY[p_col_index::text],
      to_jsonb(v_new_cell_version),
      true
    );

    UPDATE public.workbook_rows
    SET
      cells = v_cells,
      row_version = v_new_row_version,
      cell_versions = v_cell_versions
    WHERE id = v_row.id;
  ELSE
    -- New row: baseline version must be 0, and new row MUST inherit parent sheet's import_batch_id
    IF p_expected_cell_version <> 0 THEN
      RAISE EXCEPTION 'VERSION_CONFLICT: Cell (%, %) expected version %, but current version is 0.',
        p_row_index, p_col_index, p_expected_cell_version;
    END IF;

    v_new_cell_version := 1;
    v_new_row_version := 1;

    v_cells := '[]'::jsonb;
    WHILE jsonb_array_length(v_cells) <= p_col_index LOOP
      v_cells := v_cells || jsonb_build_array('');
    END LOOP;

    v_cells := jsonb_set(v_cells, ARRAY[p_col_index::text], p_new_value, false);
    v_cell_versions := jsonb_build_object(p_col_index::text, v_new_cell_version);

    INSERT INTO public.workbook_rows (
      sheet_id,
      row_index,
      cells,
      row_version,
      import_batch_id,
      cell_versions
    )
    VALUES (
      p_sheet_id,
      p_row_index,
      v_cells,
      v_new_row_version,
      v_sheet.import_batch_id,
      v_cell_versions
    );
  END IF;

  RETURN jsonb_build_object(
    'sheetId', p_sheet_id,
    'rowIndex', p_row_index,
    'colIndex', p_col_index,
    'newCellVersion', v_new_cell_version,
    'newRowVersion', v_new_row_version,
    'updatedAt', v_now
  );
END;
$$;

-- 4. Recreate patch_workspace_cells_batch with atomic batch updates & parent import_batch_id inheritance
CREATE OR REPLACE FUNCTION public.patch_workspace_cells_batch(
  p_workspace text,
  p_sheet_id uuid,
  p_edits jsonb
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
DECLARE
  v_user_id uuid;
  v_staff record;
  v_sheet record;
  v_row record;

  v_edit_count integer := 0;
  v_edit_elem jsonb;
  v_row_index integer;
  v_col_index integer;
  v_new_value jsonb;
  v_expected_cell_version bigint;

  v_current_cell_version bigint := 0;
  v_new_cell_version bigint := 1;
  v_new_row_version bigint := 1;
  v_cells jsonb;
  v_cell_versions jsonb;
  v_updated_cells jsonb := '[]'::jsonb;
  v_now timestamptz := now();
BEGIN
  -- 1. Authenticate active staff user & verify workspace permission
  v_user_id := auth.uid();
  IF v_user_id IS NULL THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Authentication required.';
  END IF;

  SELECT user_id, role::text AS role_text, is_active
  INTO v_staff
  FROM public.staff_profiles
  WHERE user_id = v_user_id;

  IF NOT FOUND OR v_staff.is_active IS NOT TRUE THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Inactive or missing staff profile.';
  END IF;

  IF p_workspace IS NULL OR p_workspace NOT IN ('reception', 'doctor') THEN
    RAISE EXCEPTION 'INVALID_WORKSPACE: p_workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % role cannot modify the % workspace.', v_staff.role_text, p_workspace;
  END IF;

  -- 2. Validate batch payload
  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_edits IS NULL OR jsonb_typeof(p_edits) IS DISTINCT FROM 'array' THEN
    RAISE EXCEPTION 'INVALID_EDITS: p_edits must be a JSON array.';
  END IF;

  v_edit_count := jsonb_array_length(p_edits);
  IF v_edit_count = 0 OR v_edit_count > 500 THEN
    RAISE EXCEPTION 'INVALID_EDITS: p_edits must contain between 1 and 500 edits.';
  END IF;

  -- 3. Lock parent worksheet and capture its exact import_batch_id
  SELECT ws.id, ws.workspace, ws.import_batch_id
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet % not found in % workspace.', p_sheet_id, p_workspace;
  END IF;

  -- 4. Apply every edit atomically in order
  FOR v_edit_elem IN
    SELECT value
    FROM jsonb_array_elements(p_edits)
  LOOP
    IF jsonb_typeof(v_edit_elem) IS DISTINCT FROM 'object' THEN
      RAISE EXCEPTION 'INVALID_EDITS: Every edit entry must be a JSON object.';
    END IF;

    IF jsonb_typeof(v_edit_elem->'rowIndex') IS DISTINCT FROM 'number'
       OR (v_edit_elem->>'rowIndex') !~ '^[0-9]+$' THEN
      RAISE EXCEPTION 'INVALID_ROW_INDEX: Edit rowIndex must be a nonnegative integer.';
    END IF;
    v_row_index := (v_edit_elem->>'rowIndex')::integer;
    IF v_row_index < 0 OR v_row_index >= 100000 THEN
      RAISE EXCEPTION 'INVALID_ROW_INDEX: rowIndex must be between 0 and 99999.';
    END IF;

    IF jsonb_typeof(v_edit_elem->'colIndex') IS DISTINCT FROM 'number'
       OR (v_edit_elem->>'colIndex') !~ '^[0-9]+$' THEN
      RAISE EXCEPTION 'INVALID_COL_INDEX: Edit colIndex must be a nonnegative integer.';
    END IF;
    v_col_index := (v_edit_elem->>'colIndex')::integer;
    IF v_col_index < 0 OR v_col_index >= 256 THEN
      RAISE EXCEPTION 'INVALID_COL_INDEX: colIndex must be between 0 and 255.';
    END IF;

    v_new_value := v_edit_elem->'value';
    IF v_new_value IS NULL OR jsonb_typeof(v_new_value) NOT IN ('string', 'number') THEN
      RAISE EXCEPTION 'INVALID_CELL_VALUE: Cell value at (%, %) must be a JSON string or number.', v_row_index, v_col_index;
    END IF;

    IF jsonb_typeof(v_edit_elem->'expectedCellVersion') IS DISTINCT FROM 'number'
       OR (v_edit_elem->>'expectedCellVersion') !~ '^[0-9]+$' THEN
      RAISE EXCEPTION 'CONCURRENCY_REQUIRED: expectedCellVersion at (%, %) must be a nonnegative integer.', v_row_index, v_col_index;
    END IF;
    v_expected_cell_version := (v_edit_elem->>'expectedCellVersion')::bigint;

    SELECT wr.id, wr.cells, wr.row_version, wr.cell_versions, wr.import_batch_id
    INTO v_row
    FROM public.workbook_rows wr
    WHERE wr.sheet_id = p_sheet_id
      AND wr.row_index = v_row_index
    FOR UPDATE;

    IF FOUND THEN
      v_current_cell_version := COALESCE((v_row.cell_versions ->> v_col_index::text)::bigint, 0);

      IF v_current_cell_version <> v_expected_cell_version THEN
        RAISE EXCEPTION 'VERSION_CONFLICT: Cell (%, %) expected version %, but current version is %.',
          v_row_index, v_col_index, v_expected_cell_version, v_current_cell_version;
      END IF;

      v_new_cell_version := v_current_cell_version + 1;
      v_new_row_version := COALESCE(v_row.row_version, 0) + 1;

      v_cells := CASE
        WHEN v_row.cells IS NOT NULL AND jsonb_typeof(v_row.cells) = 'array' THEN v_row.cells
        ELSE '[]'::jsonb
      END;

      WHILE jsonb_array_length(v_cells) <= v_col_index LOOP
        v_cells := v_cells || jsonb_build_array('');
      END LOOP;

      v_cells := jsonb_set(v_cells, ARRAY[v_col_index::text], v_new_value, false);

      v_cell_versions := CASE
        WHEN v_row.cell_versions IS NOT NULL AND jsonb_typeof(v_row.cell_versions) = 'object' THEN v_row.cell_versions
        ELSE '{}'::jsonb
      END;

      v_cell_versions := jsonb_set(
        v_cell_versions,
        ARRAY[v_col_index::text],
        to_jsonb(v_new_cell_version),
        true
      );

      UPDATE public.workbook_rows
      SET
        cells = v_cells,
        row_version = v_new_row_version,
        cell_versions = v_cell_versions
      WHERE id = v_row.id;
    ELSE
      IF v_expected_cell_version <> 0 THEN
        RAISE EXCEPTION 'VERSION_CONFLICT: Cell (%, %) expected version %, but current version is 0.',
          v_row_index, v_col_index, v_expected_cell_version;
      END IF;

      v_new_cell_version := 1;
      v_new_row_version := 1;

      v_cells := '[]'::jsonb;
      WHILE jsonb_array_length(v_cells) <= v_col_index LOOP
        v_cells := v_cells || jsonb_build_array('');
      END LOOP;

      v_cells := jsonb_set(v_cells, ARRAY[v_col_index::text], v_new_value, false);
      v_cell_versions := jsonb_build_object(v_col_index::text, v_new_cell_version);

      INSERT INTO public.workbook_rows (
        sheet_id,
        row_index,
        cells,
        row_version,
        import_batch_id,
        cell_versions
      )
      VALUES (
        p_sheet_id,
        v_row_index,
        v_cells,
        v_new_row_version,
        v_sheet.import_batch_id,
        v_cell_versions
      );
    END IF;

    v_updated_cells := v_updated_cells || jsonb_build_array(
      jsonb_build_object(
        'rowIndex', v_row_index,
        'colIndex', v_col_index,
        'newCellVersion', v_new_cell_version,
        'newRowVersion', v_new_row_version
      )
    );
  END LOOP;

  RETURN jsonb_build_object(
    'sheetId', p_sheet_id,
    'editCount', v_edit_count,
    'updatedCells', v_updated_cells,
    'updatedAt', v_now
  );
END;
$$;

REVOKE ALL ON FUNCTION public.patch_workspace_cell(text, uuid, integer, integer, jsonb, bigint) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.patch_workspace_cell(text, uuid, integer, integer, jsonb, bigint) TO authenticated;

REVOKE ALL ON FUNCTION public.patch_workspace_cells_batch(text, uuid, jsonb) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.patch_workspace_cells_batch(text, uuid, jsonb) TO authenticated;
