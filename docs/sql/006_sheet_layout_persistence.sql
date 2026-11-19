-- ============================================================================
-- MIGRATION 006: WORKSHEET COLUMN WIDTH & ROW HEIGHT LAYOUT PERSISTENCE
-- ============================================================================
-- Safe, rerunnable migration adding per-worksheet layout columns and an
-- authenticated, role-verified RPC so resized column widths and row heights
-- persist in Supabase across sessions and devices.
-- ============================================================================

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS column_widths jsonb NOT NULL DEFAULT '{}'::jsonb;

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS row_heights jsonb NOT NULL DEFAULT '{}'::jsonb;

DROP FUNCTION IF EXISTS public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb);

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
      IF v_entry.key ~ '^[0-9]{1,6}$' AND jsonb_typeof(v_entry.value) = 'number' THEN
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
      IF v_entry.key ~ '^[0-9]{1,6}$' AND jsonb_typeof(v_entry.value) = 'number' THEN
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

REVOKE ALL ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) TO authenticated;

NOTIFY pgrst, 'reload schema';
