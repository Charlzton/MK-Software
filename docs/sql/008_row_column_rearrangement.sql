-- ============================================================================
-- MIGRATION 008: ATOMIC ROW & COLUMN REARRANGEMENT RPCS (HARDENED)
-- ============================================================================
-- Safe, rerunnable migration building cleanly on Migrations 001-007.
-- Does NOT modify or overwrite previously applied migrations or seed data.
--
-- What this provides:
--   1. Idempotently ensures workbook_sheets.column_widths and row_heights exist.
--   2. Drops only the exact function signatures for reorder_workspace_rows and
--      reorder_workspace_columns (no broad pg_proc overload sweep).
--   3. public.reorder_workspace_rows(
--        p_workspace, p_sheet_id, p_start_row_index, p_end_row_index, p_target_row_index
--      )
--      Locks the parent workbook_sheets row FOR UPDATE and all workbook_rows for
--      that sheet in deterministic (row_index, id) order FOR UPDATE; validates
--      that the selected row range exists and p_target_row_index is within the
--      actual worksheet data-row range (preventing sparse out-of-range shifts);
--      atomically moves rows using a two-phase +1000000 shift and remaps
--      workbook_sheets.row_heights in lockstep.
--   4. public.reorder_workspace_columns(
--        p_workspace, p_sheet_id, p_start_col_index, p_end_col_index, p_target_col_index
--      )
--      Locks the parent workbook_sheets row FOR UPDATE and all workbook_rows for
--      that sheet in deterministic (row_index, id) order FOR UPDATE; validates
--      the selected column range and target against actual worksheet header and
--      data dimensions (preventing blank-header padding from invalid targets);
--      atomically permutes headers, header_version, column_widths, and every
--      row's cells, cell_versions, and row_version.
--   5. Reloads PostgREST schema cache.
-- ============================================================================

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS column_widths jsonb NOT NULL DEFAULT '{}'::jsonb;

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS row_heights jsonb NOT NULL DEFAULT '{}'::jsonb;

DROP FUNCTION IF EXISTS public.reorder_workspace_rows(text, uuid, integer, integer, integer);
DROP FUNCTION IF EXISTS public.reorder_workspace_columns(text, uuid, integer, integer, integer);

-- 1. Atomic Row Rearrangement RPC
CREATE OR REPLACE FUNCTION public.reorder_workspace_rows(
  p_workspace text,
  p_sheet_id uuid,
  p_start_row_index integer,
  p_end_row_index integer,
  p_target_row_index integer
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
  v_selected_row_count integer;
  v_affected_row_count integer;
  v_affected_expected_count integer;
  v_count integer;
  v_min integer;
  v_max integer;
  v_new_heights jsonb := '{}'::jsonb;
  v_entry record;
  v_grid_idx integer;
  v_data_idx integer;
  v_new_data_idx integer;
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

  IF p_start_row_index IS NULL OR p_end_row_index IS NULL OR p_target_row_index IS NULL
     OR p_start_row_index < 0 OR p_end_row_index < p_start_row_index
     OR p_end_row_index >= 100000 OR p_target_row_index < 0 OR p_target_row_index > 100000 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Invalid row reorder parameters.';
  END IF;

  -- Lock parent sheet first (canonical lock order shared with cell-patch & structural RPCs)
  SELECT ws.id, ws.row_heights
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  -- Lock all rows of this worksheet in deterministic order to prevent concurrent
  -- cell edits, row insertions, or deletions from interleaving with bounds checks or shifts.
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

  IF v_total_rows = 0 OR v_max_row_idx < 0 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Worksheet has no data rows to reorder.';
  END IF;

  IF v_min_row_idx < 0 OR v_max_row_idx >= 100000 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Existing worksheet row indexes (%..%) are outside supported bounds (0..99999).',
      v_min_row_idx, v_max_row_idx;
  END IF;

  IF p_end_row_index > v_max_row_idx THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Selected row range (%..%) exceeds worksheet data bounds (0..%).',
      p_start_row_index, p_end_row_index, v_max_row_idx;
  END IF;

  IF p_target_row_index > v_max_row_idx + 1 THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Target row position (%) exceeds valid worksheet insertion range (0..%).',
      p_target_row_index, v_max_row_idx + 1;
  END IF;

  v_count := p_end_row_index - p_start_row_index + 1;

  SELECT COUNT(DISTINCT wr.row_index)::integer
  INTO v_selected_row_count
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index BETWEEN p_start_row_index AND p_end_row_index;

  IF v_selected_row_count <> v_count THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Selected row range (%..%) contains missing row indexes (expected % rows, found %).',
      p_start_row_index, p_end_row_index, v_count, v_selected_row_count;
  END IF;

  -- Dropping inside [start .. end + 1] is a validated no-op
  IF p_target_row_index >= p_start_row_index AND p_target_row_index <= p_end_row_index + 1 THEN
    RETURN jsonb_build_object(
      'success', true,
      'sheetId', p_sheet_id,
      'rowHeights', COALESCE(v_sheet.row_heights, '{}'::jsonb)
    );
  END IF;

  IF p_target_row_index < p_start_row_index THEN
    v_min := p_target_row_index;
    v_max := p_end_row_index;
  ELSE
    v_min := p_start_row_index;
    v_max := p_target_row_index - 1;
  END IF;

  v_affected_expected_count := v_max - v_min + 1;

  SELECT COUNT(DISTINCT wr.row_index)::integer
  INTO v_affected_row_count
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
    AND wr.row_index BETWEEN v_min AND v_max;

  IF v_affected_row_count <> v_affected_expected_count THEN
    RAISE EXCEPTION 'INVALID_ROW_INDEX: Affected row interval (%..%) contains missing row indexes (expected % rows, found %).',
      v_min, v_max, v_affected_expected_count, v_affected_row_count;
  END IF;

  IF p_target_row_index < p_start_row_index THEN
    UPDATE public.workbook_rows wr
    SET row_index = wr.row_index + 1000000
    WHERE wr.sheet_id = p_sheet_id
      AND wr.row_index BETWEEN v_min AND v_max;

    UPDATE public.workbook_rows wr
    SET row_index = CASE
      WHEN (wr.row_index - 1000000) BETWEEN p_start_row_index AND p_end_row_index
        THEN (wr.row_index - 1000000) - (p_start_row_index - p_target_row_index)
      ELSE (wr.row_index - 1000000) + v_count
    END
    WHERE wr.sheet_id = p_sheet_id
      AND wr.row_index BETWEEN v_min + 1000000 AND v_max + 1000000;
  ELSE
    UPDATE public.workbook_rows wr
    SET row_index = wr.row_index + 1000000
    WHERE wr.sheet_id = p_sheet_id
      AND wr.row_index BETWEEN v_min AND v_max;

    UPDATE public.workbook_rows wr
    SET row_index = CASE
      WHEN (wr.row_index - 1000000) BETWEEN p_start_row_index AND p_end_row_index
        THEN (wr.row_index - 1000000) + (p_target_row_index - p_end_row_index - 1)
      ELSE (wr.row_index - 1000000) - v_count
    END
    WHERE wr.sheet_id = p_sheet_id
      AND wr.row_index BETWEEN v_min + 1000000 AND v_max + 1000000;
  END IF;

  -- Remap workbook_sheets.row_heights (gridRowIdx 0 is Row 1 header; data rows start at gridRowIdx = 1)
  IF v_sheet.row_heights IS NOT NULL AND jsonb_typeof(v_sheet.row_heights) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.row_heights) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_grid_idx := v_entry.key::integer;
        IF v_grid_idx = 0 THEN
          v_new_heights := v_new_heights || jsonb_build_object('0', v_entry.value);
        ELSE
          v_data_idx := v_grid_idx - 1;
          IF p_target_row_index < p_start_row_index THEN
            IF v_data_idx BETWEEN p_start_row_index AND p_end_row_index THEN
              v_new_data_idx := v_data_idx - (p_start_row_index - p_target_row_index);
            ELSIF v_data_idx BETWEEN p_target_row_index AND p_start_row_index - 1 THEN
              v_new_data_idx := v_data_idx + v_count;
            ELSE
              v_new_data_idx := v_data_idx;
            END IF;
          ELSE
            IF v_data_idx BETWEEN p_start_row_index AND p_end_row_index THEN
              v_new_data_idx := v_data_idx + (p_target_row_index - p_end_row_index - 1);
            ELSIF v_data_idx BETWEEN p_end_row_index + 1 AND p_target_row_index - 1 THEN
              v_new_data_idx := v_data_idx - v_count;
            ELSE
              v_new_data_idx := v_data_idx;
            END IF;
          END IF;

          IF v_new_data_idx >= 0 AND v_new_data_idx + 1 < 100000 THEN
            v_new_heights := v_new_heights || jsonb_build_object((v_new_data_idx + 1)::text, v_entry.value);
          END IF;
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
    'rowHeights', v_new_heights
  );
END;
$$;

-- 2. Atomic Column Rearrangement RPC
CREATE OR REPLACE FUNCTION public.reorder_workspace_columns(
  p_workspace text,
  p_sheet_id uuid,
  p_start_col_index integer,
  p_end_col_index integer,
  p_target_col_index integer
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
  v_count integer;
  v_headers jsonb;
  v_header_len integer;
  v_max_cells_len integer;
  v_effective_col_count integer;
  v_header_domain_end integer;
  v_new_headers jsonb := '[]'::jsonb;
  v_new_header_version bigint;
  v_new_widths jsonb := '{}'::jsonb;
  v_entry record;
  v_old_idx integer;
  v_new_idx integer;
  v_cells jsonb;
  v_cells_len integer;
  v_row_domain_end integer;
  v_min_affected integer;
  v_new_cells jsonb;
  v_new_cell_versions jsonb;
  v_src_idx integer;
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

  IF p_start_col_index IS NULL OR p_end_col_index IS NULL OR p_target_col_index IS NULL
     OR p_start_col_index < 0 OR p_end_col_index < p_start_col_index
     OR p_end_col_index >= 256 OR p_target_col_index < 0 OR p_target_col_index > 256 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Invalid column reorder parameters.';
  END IF;

  -- Lock parent sheet first (canonical lock order shared with cell-patch & structural RPCs)
  SELECT ws.id, ws.headers, ws.header_version, ws.column_widths
  INTO v_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  FOR UPDATE;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  -- Lock all rows of this worksheet in deterministic order before reading data dimensions
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

  IF v_header_len > 256 OR v_max_cells_len > 256 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Worksheet exceeds the supported 256-column limit (headers: %, max row cells: %).',
      v_header_len, v_max_cells_len;
  END IF;

  v_effective_col_count := LEAST(256, GREATEST(v_header_len, v_max_cells_len));

  IF v_effective_col_count <= 0 THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Worksheet has no columns to reorder.';
  END IF;

  IF p_end_col_index >= v_effective_col_count THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Selected column range (%..%) exceeds worksheet column bounds (0..%).',
      p_start_col_index, p_end_col_index, v_effective_col_count - 1;
  END IF;

  IF p_target_col_index > v_effective_col_count THEN
    RAISE EXCEPTION 'INVALID_COL_INDEX: Target column position (%) exceeds valid worksheet insertion range (0..%).',
      p_target_col_index, v_effective_col_count;
  END IF;

  -- Dropping inside [start .. end + 1] is a validated no-op
  IF p_target_col_index >= p_start_col_index AND p_target_col_index <= p_end_col_index + 1 THEN
    RETURN jsonb_build_object(
      'success', true,
      'sheetId', p_sheet_id,
      'headers', v_headers,
      'headerVersion', COALESCE(v_sheet.header_version, 1),
      'columnWidths', COALESCE(v_sheet.column_widths, '{}'::jsonb)
    );
  END IF;

  v_count := p_end_col_index - p_start_col_index + 1;
  v_min_affected := LEAST(p_start_col_index, p_target_col_index);
  v_header_domain_end := LEAST(255, GREATEST(v_header_len - 1, p_end_col_index, p_target_col_index - 1));

  -- Permute headers array
  FOR v_new_idx IN 0..v_header_domain_end LOOP
    IF p_target_col_index < p_start_col_index THEN
      IF v_new_idx < p_target_col_index THEN
        v_src_idx := v_new_idx;
      ELSIF v_new_idx < p_target_col_index + v_count THEN
        v_src_idx := p_start_col_index + (v_new_idx - p_target_col_index);
      ELSIF v_new_idx <= p_end_col_index THEN
        v_src_idx := v_new_idx - v_count;
      ELSE
        v_src_idx := v_new_idx;
      END IF;
    ELSE
      IF v_new_idx < p_start_col_index THEN
        v_src_idx := v_new_idx;
      ELSIF v_new_idx < p_target_col_index - v_count THEN
        v_src_idx := v_new_idx + v_count;
      ELSIF v_new_idx < p_target_col_index THEN
        v_src_idx := p_start_col_index + (v_new_idx - (p_target_col_index - v_count));
      ELSE
        v_src_idx := v_new_idx;
      END IF;
    END IF;

    IF v_src_idx < v_header_len THEN
      v_new_headers := v_new_headers || jsonb_build_array(v_headers -> v_src_idx);
    ELSE
      v_new_headers := v_new_headers || jsonb_build_array('');
    END IF;
  END LOOP;

  -- Permute column_widths map keys
  IF v_sheet.column_widths IS NOT NULL AND jsonb_typeof(v_sheet.column_widths) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(v_sheet.column_widths) LOOP
      IF v_entry.key ~ '^[0-9]{1,6}$' THEN
        v_old_idx := v_entry.key::integer;
        IF p_target_col_index < p_start_col_index THEN
          IF v_old_idx BETWEEN p_start_col_index AND p_end_col_index THEN
            v_new_idx := v_old_idx - (p_start_col_index - p_target_col_index);
          ELSIF v_old_idx BETWEEN p_target_col_index AND p_start_col_index - 1 THEN
            v_new_idx := v_old_idx + v_count;
          ELSE
            v_new_idx := v_old_idx;
          END IF;
        ELSE
          IF v_old_idx BETWEEN p_start_col_index AND p_end_col_index THEN
            v_new_idx := v_old_idx + (p_target_col_index - p_end_col_index - 1);
          ELSIF v_old_idx BETWEEN p_end_col_index + 1 AND p_target_col_index - 1 THEN
            v_new_idx := v_old_idx - v_count;
          ELSE
            v_new_idx := v_old_idx;
          END IF;
        END IF;

        IF v_new_idx >= 0 AND v_new_idx < 256 THEN
          v_new_widths := v_new_widths || jsonb_build_object(v_new_idx::text, v_entry.value);
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

  -- Permute cells & cell_versions across all locked rows in deterministic order
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

    IF v_cells_len > v_min_affected THEN
      v_row_domain_end := LEAST(255, GREATEST(v_cells_len - 1, p_end_col_index, p_target_col_index - 1));
      v_new_cells := '[]'::jsonb;
      FOR v_new_idx IN 0..v_row_domain_end LOOP
        IF p_target_col_index < p_start_col_index THEN
          IF v_new_idx < p_target_col_index THEN
            v_src_idx := v_new_idx;
          ELSIF v_new_idx < p_target_col_index + v_count THEN
            v_src_idx := p_start_col_index + (v_new_idx - p_target_col_index);
          ELSIF v_new_idx <= p_end_col_index THEN
            v_src_idx := v_new_idx - v_count;
          ELSE
            v_src_idx := v_new_idx;
          END IF;
        ELSE
          IF v_new_idx < p_start_col_index THEN
            v_src_idx := v_new_idx;
          ELSIF v_new_idx < p_target_col_index - v_count THEN
            v_src_idx := v_new_idx + v_count;
          ELSIF v_new_idx < p_target_col_index THEN
            v_src_idx := p_start_col_index + (v_new_idx - (p_target_col_index - v_count));
          ELSE
            v_src_idx := v_new_idx;
          END IF;
        END IF;

        IF v_src_idx < v_cells_len THEN
          v_new_cells := v_new_cells || jsonb_build_array(v_cells -> v_src_idx);
        ELSE
          v_new_cells := v_new_cells || jsonb_build_array('');
        END IF;
      END LOOP;
    ELSE
      v_new_cells := v_cells;
    END IF;

    v_new_cell_versions := '{}'::jsonb;
    IF v_row.cell_versions IS NOT NULL AND jsonb_typeof(v_row.cell_versions) = 'object' THEN
      FOR v_entry IN SELECT key, value FROM jsonb_each(v_row.cell_versions) LOOP
        IF v_entry.key ~ '^[0-9]{1,6}$' THEN
          v_old_idx := v_entry.key::integer;
          IF p_target_col_index < p_start_col_index THEN
            IF v_old_idx BETWEEN p_start_col_index AND p_end_col_index THEN
              v_new_idx := v_old_idx - (p_start_col_index - p_target_col_index);
            ELSIF v_old_idx BETWEEN p_target_col_index AND p_start_col_index - 1 THEN
              v_new_idx := v_old_idx + v_count;
            ELSE
              v_new_idx := v_old_idx;
            END IF;
          ELSE
            IF v_old_idx BETWEEN p_start_col_index AND p_end_col_index THEN
              v_new_idx := v_old_idx + (p_target_col_index - p_end_col_index - 1);
            ELSIF v_old_idx BETWEEN p_end_col_index + 1 AND p_target_col_index - 1 THEN
              v_new_idx := v_old_idx - v_count;
            ELSE
              v_new_idx := v_old_idx;
            END IF;
          END IF;

          IF v_new_idx >= 0 AND v_new_idx < 256 THEN
            v_new_cell_versions := v_new_cell_versions || jsonb_build_object(v_new_idx::text, v_entry.value);
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
    'headers', v_new_headers,
    'headerVersion', v_new_header_version,
    'columnWidths', v_new_widths
  );
END;
$$;

REVOKE ALL ON FUNCTION public.reorder_workspace_rows(text, uuid, integer, integer, integer) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.reorder_workspace_columns(text, uuid, integer, integer, integer) FROM PUBLIC, anon;

GRANT EXECUTE ON FUNCTION public.reorder_workspace_rows(text, uuid, integer, integer, integer) TO authenticated;
GRANT EXECUTE ON FUNCTION public.reorder_workspace_columns(text, uuid, integer, integer, integer) TO authenticated;

NOTIFY pgrst, 'reload schema';
