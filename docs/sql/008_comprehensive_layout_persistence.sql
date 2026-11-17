-- ============================================================================
-- MIGRATION 008: BULLETPROOF WORKSPACE SHEET LAYOUT PERSISTENCE
-- ============================================================================
-- Safe, idempotent, rerunnable migration for MK Eye Center Workbook System.
-- Corrects all column width and row height layout persistence failures in
-- Supabase without modifying existing clinical patient records or credentials.
--
-- How to apply:
--   1. Open your Supabase Dashboard: https://supabase.com/dashboard/project/_/sql
--   2. Paste this entire file into the SQL Editor.
--   3. Click Run.
-- ============================================================================

-- 1. Ensure columns exist on public.workbook_sheets with NOT NULL and default '{}'
ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS column_widths jsonb NOT NULL DEFAULT '{}'::jsonb;

ALTER TABLE public.workbook_sheets
  ADD COLUMN IF NOT EXISTS row_heights jsonb NOT NULL DEFAULT '{}'::jsonb;

UPDATE public.workbook_sheets
SET
  column_widths = COALESCE(column_widths, '{}'::jsonb),
  row_heights = COALESCE(row_heights, '{}'::jsonb)
WHERE column_widths IS NULL OR row_heights IS NULL;

-- 2. Cleanly drop any prior ambiguous overloads of save_workspace_sheet_layout
DO $$
DECLARE
  v_sig text;
BEGIN
  FOR v_sig IN
    SELECT p.oid::regprocedure::text
    FROM pg_proc p
    JOIN pg_namespace n ON n.oid = p.pronamespace
    WHERE n.nspname = 'public'
      AND p.proname = 'save_workspace_sheet_layout'
  LOOP
    EXECUTE 'DROP FUNCTION IF EXISTS ' || v_sig || ';';
  END LOOP;
END;
$$;

-- 3. Create canonical layout persistence function accepting p_sheet_id as text
CREATE OR REPLACE FUNCTION public.save_workspace_sheet_layout(
  p_workspace text,
  p_sheet_id text,
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
  v_target_uuid uuid;
  v_workspace_lower text;
  v_role_lower text;
  v_clean_cols jsonb := '{}'::jsonb;
  v_clean_rows jsonb := '{}'::jsonb;
  v_entry record;
  v_idx integer;
  v_num numeric;
  v_raw_text text;
BEGIN
  -- Authenticate calling user
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

  v_workspace_lower := LOWER(TRIM(p_workspace));
  IF v_workspace_lower IS NULL OR v_workspace_lower NOT IN ('reception', 'doctor') THEN
    RAISE EXCEPTION 'INVALID_WORKSPACE: Workspace must be reception or doctor.';
  END IF;

  v_role_lower := LOWER(TRIM(v_staff.role_text));
  IF NOT (
    v_role_lower = 'admin'
    OR (v_role_lower = 'reception' AND v_workspace_lower = 'reception')
    OR (v_role_lower = 'doctor' AND v_workspace_lower = 'doctor')
  ) THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: % cannot modify % workspace.', v_staff.role_text, p_workspace;
  END IF;

  IF p_sheet_id IS NULL OR TRIM(p_sheet_id) = '' THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: p_sheet_id is required.';
  END IF;

  BEGIN
    v_target_uuid := TRIM(p_sheet_id)::uuid;
  EXCEPTION WHEN OTHERS THEN
    RAISE EXCEPTION 'INVALID_SHEET_ID: % is not a valid UUID.', p_sheet_id;
  END;

  -- Clean & validate column widths (0..255, 40..800 px)
  IF p_column_widths IS NOT NULL AND jsonb_typeof(p_column_widths) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(p_column_widths) LOOP
      IF v_entry.key ~ '^[0-9]+$' THEN
        v_raw_text := btrim(v_entry.value #>> '{}');
        IF jsonb_typeof(v_entry.value) IN ('number', 'string') AND v_raw_text ~ '^[0-9]+(\.[0-9]+)?$' THEN
          v_idx := v_entry.key::integer;
          v_num := round(v_raw_text::numeric);
          IF v_idx >= 0 AND v_idx < 256 AND v_num >= 40 AND v_num <= 800 THEN
            v_clean_cols := v_clean_cols || jsonb_build_object(v_entry.key, v_num);
          END IF;
        END IF;
      END IF;
    END LOOP;
  END IF;

  -- Clean & validate row heights (0..99999, 24..600 px)
  IF p_row_heights IS NOT NULL AND jsonb_typeof(p_row_heights) = 'object' THEN
    FOR v_entry IN SELECT key, value FROM jsonb_each(p_row_heights) LOOP
      IF v_entry.key ~ '^[0-9]+$' THEN
        v_raw_text := btrim(v_entry.value #>> '{}');
        IF jsonb_typeof(v_entry.value) IN ('number', 'string') AND v_raw_text ~ '^[0-9]+(\.[0-9]+)?$' THEN
          v_idx := v_entry.key::integer;
          v_num := round(v_raw_text::numeric);
          IF v_idx >= 0 AND v_idx < 100000 AND v_num >= 24 AND v_num <= 600 THEN
            v_clean_rows := v_clean_rows || jsonb_build_object(v_entry.key, v_num);
          END IF;
        END IF;
      END IF;
    END LOOP;
  END IF;

  UPDATE public.workbook_sheets ws
  SET column_widths = v_clean_cols,
      row_heights = v_clean_rows
  WHERE ws.id = v_target_uuid
    AND ws.workspace = v_workspace_lower::public.workspace_type
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

-- 4. Overload supporting direct uuid callers for PostgREST signature matching
CREATE OR REPLACE FUNCTION public.save_workspace_sheet_layout(
  p_workspace text,
  p_sheet_id uuid,
  p_column_widths jsonb DEFAULT '{}'::jsonb,
  p_row_heights jsonb DEFAULT '{}'::jsonb
)
RETURNS jsonb
LANGUAGE sql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
  SELECT public.save_workspace_sheet_layout(
    p_workspace,
    p_sheet_id::text,
    p_column_widths,
    p_row_heights
  );
$$;

-- 5. Set permissions and reload schema
REVOKE ALL ON FUNCTION public.save_workspace_sheet_layout(text, text, jsonb, jsonb) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) FROM PUBLIC, anon;

GRANT EXECUTE ON FUNCTION public.save_workspace_sheet_layout(text, text, jsonb, jsonb) TO authenticated, service_role;
GRANT EXECUTE ON FUNCTION public.save_workspace_sheet_layout(text, uuid, jsonb, jsonb) TO authenticated, service_role;

-- Notify PostgREST to immediately refresh its schema cache
NOTIFY pgrst, 'reload schema';
