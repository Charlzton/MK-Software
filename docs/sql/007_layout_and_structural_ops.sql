-- ============================================================================
-- MIGRATION 007: LAYOUT PERSISTENCE REPAIR & ATOMIC ROW/COLUMN OPERATIONS
-- ============================================================================
-- Safe, rerunnable migration building cleanly on Migrations 001-006.
-- Does NOT modify or overwrite previously applied migrations or seed data.
--
-- What this provides:
--   1. Idempotently guarantees public.workbook_sheets.column_widths and
--      row_heights exist with NOT NULL DEFAULT '{}'::jsonb and repairs any NULLs.
--   2. Recreates public.save_workspace_sheet_layout cleanly (dropping only the
--      exact signature) with hardened jsonb key & numeric coercion.
--   3. Adds transactional, role-verified structural RPCs using canonical lock
--      ordering (workbook_sheets FOR UPDATE -> workbook_rows ORDER BY row_index, id FOR UPDATE):
--      - public.insert_workspace_row(p_workspace, p_sheet_id, p_row_index)
--      - public.delete_workspace_row(p_workspace, p_sheet_id, p_start_row_index, p_end_row_index)
--      - public.insert_workspace_column(p_workspace, p_sheet_id, p_col_index)
--      - public.delete_workspace_column(p_workspace, p_sheet_id, p_col_index)
--      Each RPC preserves parent import_batch_id, row/cell/header versions,
--      and shifts per-sheet column_widths / row_heights cleanly without silent
--      data truncation or unique-index collisions.
--   4. Reloads PostgREST schema cache.
-- ============================================================================

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS column_widths jsonb NOT NULL DEFAULT '{}'::jsonb;

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS row_heights jsonb NOT NULL DEFAULT '{}'::jsonb;

UPDATE public.workbook_sheets
SET
  column_widths = COALESCE(column_widths, '{}'::jsonb),
  row_heights = COALESCE(row_heights, '{}'::jsonb)
WHERE column_widths IS NULL OR row_heights IS NULL;

DROP FUNCTION IF EXISTS public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb);
DROP FUNCTION IF EXISTS public.insert_workspace_row(text, uuid, integer);
DROP FUNCTION IF EXISTS public.delete_workspace_row(text, uuid, integer, integer);
DROP FUNCTION IF EXISTS public.insert_workspace_column(text, uuid, integer);
DROP FUNCTION IF EXISTS public.delete_workspace_column(text, uuid, integer);

-- 1. Hardened Layout Persistence RPC (Column Widths & Row Heights)
CREATE OR REPLACE FUNCTION public.save_workspace_sheet_layout(
  p_workspace text,
  p_sheet_id uuid,
  p_column_widths jsonb DEFAULT '{}'::jsonb,
  p_row_heights jsonb DEFAULT '{}'::jsonb
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
  v_clean_cols jsonb := '{}'::jsonb;
  v_clean_rows jsonb := '{}'::jsonb;
  v_entry record;
  v_idx integer;
  v_num numeric;
  v_raw_text text;
BEGIN
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
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_column_widths IS NOT NULL AND jsonb_typeof(p_column_widths) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(p_column_widths) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' AND jsonb_typeof(v_entry.value) IN ('number', 'string') THEN
        v_raw_text := btrim(v_entry.value #>> '{}');
        IF length(v_raw_text) <= 32 AND v_raw_text ~ '^[0-9]{1,6}(\.[0-9]+)?$' THEN
          v_idx := v_entry.key::integer;
          v_num := round(v_raw_text::numeric);
          IF v_idx >= 0 AND v_idx < 256 AND v_num >= 40 AND v_num <= 800 THEN
            v_clean_cols := v_clean_cols || jsonb_build_object(v_idx::text, v_num);
          END IF;
        END IF;
      END IF;
    END LOOP;
  END IF;

  IF p_row_heights IS NOT NULL AND jsonb_typeof(p_row_heights) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(p_row_heights) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' AND jsonb_typeof(v_entry.value) IN ('number', 'string') THEN
        v_raw_text := btrim(v_entry.value #>> '{}');
        IF length(v_raw_text) <= 32 AND v_raw_text ~ '^[0-9]{1,6}(\.[0-9]+)?$' THEN
          v_idx := v_entry.key::integer;
          v_num := round(v_raw_text::numeric);
          IF v_idx >= 0 AND v_idx < 100000 AND v_num >= 24 AND v_num <= 600 THEN
            v_clean_rows := v_clean_rows || jsonb_build_object(v_idx::text, v_num);
          END IF;
        END IF;
      END IF;
    END LOOP;
  END IF;

  UPDATE public.workbook_sheets ws
  SET column_widths = v_clean_cols,
      row_heights = v_clean_rows
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  RETURNING ws.id, ws.column_widths, ws.row_heights INTO v_sheet;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', v_sheet.id,
    'columnWidths', COALESCE(v_sheet.column_widths, '{}'::jsonb),
    'rowHeights', COALESCE(v_sheet.row_heights, '{}'::jsonb)
  );
END;
$$;

-- 2. Insert Data Row RPC (shifts data rows >= p_row_index down safely)
CREATE OR REPLACE FUNCTION public.insert_workspace_row(
  p_workspace text,
  p_sheet_id uuid,
  p_row_index integer
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
  v_min_row_idx integer;
  v_max_row_idx integer;
  v_total_rows integer;
  v_new_heights jsonb := '{}'::jsonb;
  v_entry record;
  v_key_idx integer;
  v_grid_target integer;
BEGIN
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
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_row_index IS NULL OR p_row_index < 0 OR p_row_index >= 99999 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: p_row_index must be between 0 and 99998.';
  END IF;

  -- Canonical lock order: parent sheet first, then all rows in deterministic order
  SELECT ws.id, ws.row_heights
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  PERFORM wr.id
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
  ORDER BY wr.row_index, wr.id
  FOR UPDATE;

  SELECT
    COALESCE(MIN(wr.row_index), -1),
    COALESCE(MAX(wr.row_index), -1),
    COUNT(*)::integer
  INTO v_min_row_idx, v_max_row_idx, v_total_rows
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id;

  IF v_total_rows > 0 AND v_min_row_idx < 0 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Existing worksheet contains negative row_index (%).', v_min_row_idx;
  END IF;

  IF v_max_row_idx >= 99999 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Maximum row limit (100000) reached.';
  END IF;

  -- Two-phase shift avoids unique (sheet_id, row_index) collisions
  UPDATE public.workbook_rows wr
  SET row_index = wr.row_index + 1000000
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index >= p_row_index;

  UPDATE public.workbook_rows wr
  SET row_index = wr.row_index - 999999
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index >= p_row_index + 1000000;

  -- Shift manual row_heights (gridRowIdx = dataRowIdx + 1; gridRowIdx 0 header stays unchanged)
  v_grid_target := p_row_index + 1;
  IF v_sheet.row_heights IS NOT NULL AND jsonb_typeof(v_sheet.row_heights) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.row_heights) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_key_idx := v_entry.key::integer;
        IF v_key_idx >= v_grid_target THEN
          IF v_key_idx + 1 < 100000 THEN
            v_new_heights := v_new_heights || jsonb_build_object((v_key_idx + 1)::text, v_entry.value);
          END IF;
        ELSIF v_key_idx >= 0 AND v_key_idx < 100000 THEN
          v_new_heights := v_new_heights || jsonb_build_object(v_key_idx::text, v_entry.value);
        END IF;
      END IF;
    END LOOP;

    UPDATE public.workbook_sheets ws
    SET row_heights = v_new_heights
    WHERE ws.id = p_sheet_id;
  END IF;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', p_sheet_id,
    'insertedRowIndex', p_row_index,
    'rowHeights', v_new_heights
  );
END;
$$;

-- 3. Delete Data Row(s) RPC (removes range and shifts remaining rows up safely)
CREATE OR REPLACE FUNCTION public.delete_workspace_row(
  p_workspace text,
  p_sheet_id uuid,
  p_start_row_index integer,
  p_end_row_index integer DEFAULT NULL
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
  v_min_row_idx integer;
  v_max_row_idx integer;
  v_total_rows integer;
  v_end_idx integer;
  v_delete_count integer;
  v_new_heights jsonb := '{}'::jsonb;
  v_entry record;
  v_key_idx integer;
  v_grid_start integer;
  v_grid_end integer;
BEGIN
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
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  v_end_idx := COALESCE(p_end_row_index, p_start_row_index);
  IF p_start_row_index IS NULL OR p_start_row_index < 0 OR v_end_idx < p_start_row_index OR v_end_idx >= 100000 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Invalid row deletion range.';
  END IF;

  v_delete_count := v_end_idx - p_start_row_index + 1;

  -- Canonical lock order: parent sheet first, then all rows in deterministic order
  SELECT ws.id, ws.row_heights
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  PERFORM wr.id
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
  ORDER BY wr.row_index, wr.id
  FOR UPDATE;

  SELECT
    COALESCE(MIN(wr.row_index), -1),
    COALESCE(MAX(wr.row_index), -1),
    COUNT(*)::integer
  INTO v_min_row_idx, v_max_row_idx, v_total_rows
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id;

  IF v_total_rows > 0 AND (v_min_row_idx < 0 OR v_max_row_idx >= 100000) THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Existing worksheet row indexes (%..%) are outside supported bounds (0..99999).',
      v_min_row_idx, v_max_row_idx;
  END IF;

  DELETE FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index BETWEEN p_start_row_index AND v_end_idx;

  -- Two-phase downward index shift
  UPDATE public.workbook_rows wr
  SET row_index = wr.row_index + 1000000
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index > v_end_idx;

  UPDATE public.workbook_rows wr
  SET row_index = wr.row_index - (1000000 + v_delete_count)
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index > v_end_idx + 1000000;

  -- Shift manual row_heights (gridRowIdx = dataRowIdx + 1)
  v_grid_start := p_start_row_index + 1;
  v_grid_end := v_end_idx + 1;
  IF v_sheet.row_heights IS NOT NULL AND jsonb_typeof(v_sheet.row_heights) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.row_heights) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_key_idx := v_entry.key::integer;
        IF v_key_idx >= 0 AND v_key_idx < v_grid_start THEN
          v_new_heights := v_new_heights || jsonb_build_object(v_key_idx::text, v_entry.value);
        ELSIF v_key_idx > v_grid_end AND v_key_idx < 100000 THEN
          v_new_heights := v_new_heights || jsonb_build_object((v_key_idx - v_delete_count)::text, v_entry.value);
        END IF;
      END IF;
    END LOOP;

    UPDATE public.workbook_sheets ws
    SET row_heights = v_new_heights
    WHERE ws.id = p_sheet_id;
  END IF;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', p_sheet_id,
    'startRowIndex', p_start_row_index,
    'endRowIndex', v_end_idx,
    'rowHeights', v_new_heights
  );
END;
$$;

-- 4. Insert Column RPC (inserts blank column at p_col_index across headers, rows, cell_versions, widths)
CREATE OR REPLACE FUNCTION public.insert_workspace_column(
  p_workspace text,
  p_sheet_id uuid,
  p_col_index integer
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
  v_headers jsonb;
  v_new_headers jsonb := '[]'::jsonb;
  v_header_len integer;
  v_max_cells_len integer;
  v_new_header_version bigint;
  v_new_widths jsonb := '{}'::jsonb;
  v_entry record;
  v_key_idx integer;
  v_cells jsonb;
  v_new_cells jsonb;
  v_cells_len integer;
  v_new_cell_versions jsonb;
  i integer;
BEGIN
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
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_col_index IS NULL OR p_col_index < 0 OR p_col_index >= 255 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: p_col_index must be between 0 and 254.';
  END IF;

  -- Canonical lock order: parent sheet first, then all rows in deterministic order
  SELECT ws.id, ws.headers, ws.header_version, ws.column_widths
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  PERFORM wr.id
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
  ORDER BY wr.row_index, wr.id
  FOR UPDATE;

  v_headers := CASE
    WHEN v_sheet.headers IS NOT NULL AND jsonb_typeof(v_sheet.headers) = 'array' THEN v_sheet.headers
    ELSE '[]'::jsonb
  END;
  v_header_len := jsonb_array_length(v_headers);

  SELECT COALESCE(MAX(jsonb_array_length(wr.cells)), 0)
  INTO v_max_cells_len
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
    AND wr.cells IS NOT NULL
    AND jsonb_typeof(wr.cells) = 'array';

  IF v_header_len >= 256 OR v_max_cells_len >= 256 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Maximum column limit (256) reached; inserting a column would truncate column 255.';
  END IF;

  -- Build spliced headers array
  FOR i IN 0..GREATEST(v_header_len, p_col_index) LOOP
    IF i = p_col_index THEN
      v_new_headers := v_new_headers || jsonb_build_array('');
    END IF;
    IF i < v_header_len THEN
      v_new_headers := v_new_headers || jsonb_build_array(v_headers -> i);
    ELSIF i < p_col_index THEN
      v_new_headers := v_new_headers || jsonb_build_array('');
    END IF;
  END LOOP;

  -- Shift column_widths
  IF v_sheet.column_widths IS NOT NULL AND jsonb_typeof(v_sheet.column_widths) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.column_widths) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_key_idx := v_entry.key::integer;
        IF v_key_idx >= p_col_index THEN
          IF v_key_idx + 1 < 256 THEN
            v_new_widths := v_new_widths || jsonb_build_object((v_key_idx + 1)::text, v_entry.value);
          END IF;
        ELSIF v_key_idx >= 0 AND v_key_idx < 256 THEN
          v_new_widths := v_new_widths || jsonb_build_object(v_key_idx::text, v_entry.value);
        END IF;
      END IF;
    END LOOP;
  END IF;

  v_new_header_version := COALESCE(v_sheet.header_version, 1) + 1;

  UPDATE public.workbook_sheets ws
  SET headers = v_new_headers,
      header_version = v_new_header_version,
      column_widths = v_new_widths
  WHERE ws.id = p_sheet_id;

  -- Shift cells & cell_versions across all locked rows in deterministic order
  FOR v_row IN
    SELECT wr.id, wr.cells, wr.row_version, wr.cell_versions
    FROM public.workbook_rows wr
    WHERE wr.sheet_id = p_sheet_id
    ORDER BY wr.row_index, wr.id
  LOOP
    v_cells := CASE
      WHEN v_row.cells IS NOT NULL AND jsonb_typeof(v_row.cells) = 'array' THEN v_row.cells
      ELSE '[]'::jsonb
    END;
    v_cells_len := jsonb_array_length(v_cells);

    IF v_cells_len > p_col_index THEN
      v_new_cells := '[]'::jsonb;
      FOR i IN 0..(v_cells_len - 1) LOOP
        IF i = p_col_index THEN
          v_new_cells := v_new_cells || jsonb_build_array('');
        END IF;
        v_new_cells := v_new_cells || jsonb_build_array(v_cells -> i);
      END LOOP;
    ELSE
      v_new_cells := v_cells;
    END IF;

    v_new_cell_versions := '{}'::jsonb;
    IF v_row.cell_versions IS NOT NULL AND jsonb_typeof(v_row.cell_versions) = 'object' THEN
      FOR v_entry IN SELECT key, value FROM jsonb_each(v_row.cell_versions) LOOP
        IF v_entry.key ~ '^[0-9]{1,6}$' THEN
          v_key_idx := v_entry.key::integer;
          IF v_key_idx >= p_col_index THEN
            IF v_key_idx + 1 < 256 THEN
              v_new_cell_versions := v_new_cell_versions || jsonb_build_object((v_key_idx + 1)::text, v_entry.value);
            END IF;
          ELSIF v_key_idx >= 0 AND v_key_idx < 256 THEN
            v_new_cell_versions := v_new_cell_versions || jsonb_build_object(v_key_idx::text, v_entry.value);
          END IF;
        END IF;
      END LOOP;
    END IF;

    UPDATE public.workbook_rows wr
    SET cells = v_new_cells,
        row_version = COALESCE(v_row.row_version, 1) + 1,
        cell_versions = v_new_cell_versions
    WHERE wr.id = v_row.id;
  END LOOP;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', p_sheet_id,
    'insertedColIndex', p_col_index,
    'headers', v_new_headers,
    'headerVersion', v_new_header_version,
    'columnWidths', v_new_widths
  );
END;
$$;

-- 5. Delete Column RPC (removes p_col_index across headers, rows, cell_versions, widths)
CREATE OR REPLACE FUNCTION public.delete_workspace_column(
  p_workspace text,
  p_sheet_id uuid,
  p_col_index integer
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
  v_headers jsonb;
  v_new_headers jsonb := '[]'::jsonb;
  v_header_len integer;
  v_new_header_version bigint;
  v_new_widths jsonb := '{}'::jsonb;
  v_entry record;
  v_key_idx integer;
  v_cells jsonb;
  v_new_cells jsonb;
  v_cells_len integer;
  v_new_cell_versions jsonb;
  i integer;
BEGIN
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
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  IF NOT (
    v_staff.role_text = 'admin'
    OR (v_staff.role_text = 'reception' AND p_workspace = 'reception')
    OR (v_staff.role_text = 'doctor' AND p_workspace = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  IF p_col_index IS NULL OR p_col_index < 0 OR p_col_index >= 256 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: p_col_index must be between 0 and 255.';
  END IF;

  -- Canonical lock order: parent sheet first, then all rows in deterministic order
  SELECT ws.id, ws.headers, ws.header_version, ws.column_widths
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  PERFORM wr.id
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
  ORDER BY wr.row_index, wr.id
  FOR UPDATE;

  v_headers := CASE
    WHEN v_sheet.headers IS NOT NULL AND jsonb_typeof(v_sheet.headers) = 'array' THEN v_sheet.headers
    ELSE '[]'::jsonb
  END;
  v_header_len := jsonb_array_length(v_headers);

  IF v_header_len <= 1 AND p_col_index = 0 THEN
    RAISE EXCEPTION 'BAD_REQUEST: Cannot delete the only remaining column header.';
  END IF;

  FOR i IN 0..(v_header_len - 1) LOOP
    IF i <> p_col_index THEN
      v_new_headers := v_new_headers || jsonb_build_array(v_headers -> i);
    END IF;
  END LOOP;

  -- Shift column_widths
  IF v_sheet.column_widths IS NOT NULL AND jsonb_typeof(v_sheet.column_widths) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.column_widths) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_key_idx := v_entry.key::integer;
        IF v_key_idx >= 0 AND v_key_idx < p_col_index THEN
          v_new_widths := v_new_widths || jsonb_build_object(v_key_idx::text, v_entry.value);
        ELSIF v_key_idx > p_col_index AND v_key_idx < 256 THEN
          v_new_widths := v_new_widths || jsonb_build_object((v_key_idx - 1)::text, v_entry.value);
        END IF;
      END IF;
    END LOOP;
  END IF;

  v_new_header_version := COALESCE(v_sheet.header_version, 1) + 1;

  UPDATE public.workbook_sheets ws
  SET headers = v_new_headers,
      header_version = v_new_header_version,
      column_widths = v_new_widths
  WHERE ws.id = p_sheet_id;

  -- Shift cells & cell_versions across all locked rows in deterministic order
  FOR v_row IN
    SELECT wr.id, wr.cells, wr.row_version, wr.cell_versions
    FROM public.workbook_rows wr
    WHERE wr.sheet_id = p_sheet_id
    ORDER BY wr.row_index, wr.id
  LOOP
    v_cells := CASE
      WHEN v_row.cells IS NOT NULL AND jsonb_typeof(v_row.cells) = 'array' THEN v_row.cells
      ELSE '[]'::jsonb
    END;
    v_cells_len := jsonb_array_length(v_cells);

    IF v_cells_len > p_col_index THEN
      v_new_cells := '[]'::jsonb;
      FOR i IN 0..(v_cells_len - 1) LOOP
        IF i <> p_col_index THEN
          v_new_cells := v_new_cells || jsonb_build_array(v_cells -> i);
        END IF;
      END LOOP;
    ELSE
      v_new_cells := v_cells;
    END IF;

    v_new_cell_versions := '{}'::jsonb;
    IF v_row.cell_versions IS NOT NULL AND jsonb_typeof(v_row.cell_versions) = 'object' THEN
      FOR v_entry IN SELECT key, value FROM jsonb_each(v_row.cell_versions) LOOP
        IF v_entry.key ~ '^[0-9]{1,6}$' THEN
          v_key_idx := v_entry.key::integer;
          IF v_key_idx >= 0 AND v_key_idx < p_col_index THEN
            v_new_cell_versions := v_new_cell_versions || jsonb_build_object(v_key_idx::text, v_entry.value);
          ELSIF v_key_idx > p_col_index AND v_key_idx < 256 THEN
            v_new_cell_versions := v_new_cell_versions || jsonb_build_object((v_key_idx - 1)::text, v_entry.value);
          END IF;
        END IF;
      END LOOP;
    END IF;

    UPDATE public.workbook_rows wr
    SET cells = v_new_cells,
        row_version = COALESCE(v_row.row_version, 1) + 1,
        cell_versions = v_new_cell_versions
    WHERE wr.id = v_row.id;
  END LOOP;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', p_sheet_id,
    'deletedColIndex', p_col_index,
    'headers', v_new_headers,
    'headerVersion', v_new_header_version,
    'columnWidths', v_new_widths
  );
END;
$$;

REVOKE ALL ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.insert_workspace_row(text, uuid, integer) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.delete_workspace_row(text, uuid, integer, integer) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.insert_workspace_column(text, uuid, integer) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.delete_workspace_column(text, uuid, integer) FROM PUBLIC, anon;

GRANT EXECUTE ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) TO authenticated;
GRANT EXECUTE ON FUNCTION public.insert_workspace_row(text, uuid, integer) TO authenticated;
GRANT EXECUTE ON FUNCTION public.delete_workspace_row(text, uuid, integer, integer) TO authenticated;
GRANT EXECUTE ON FUNCTION public.insert_workspace_column(text, uuid, integer) TO authenticated;
GRANT EXECUTE ON FUNCTION public.delete_workspace_column(text, uuid, integer) TO authenticated;

NOTIFY pgrst, 'reload schema';
