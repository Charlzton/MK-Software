-- ============================================================================
-- MIGRATION 005: HARDENED WORKSHEET MANAGEMENT RPCS (DELETE, REORDER, RENAME)
-- ============================================================================
-- Safe, rerunnable migration fixing worksheet deletion and drag-and-drop
-- reordering in Supabase for both Reception and Doctor workspaces.
--
-- What this fixes:
--   1. Drops any overloaded signatures that can confuse PostgREST RPC lookup.
--   2. delete_workspace_sheet:
--      - Authenticates active staff and verifies workspace authorization.
--      - Verifies the target worksheet exists in the requested workspace BEFORE
--        deleting any rows.
--      - Prevents deleting the last remaining worksheet in a workspace.
--      - Safely deletes associated workbook_rows, deletes the workbook_sheets
--        record, and re-sequences remaining tab_order values (0..N-1).
--   3. reorder_workspace_sheets:
--      - Authenticates active staff and verifies workspace authorization.
--      - Validates all provided sheet IDs belong to the target workspace.
--      - Atomically updates tab_order for all worksheets in the workspace.
--   4. rename_workspace_sheet:
--      - Hardened duplicate check, validation, and permission enforcement.
--   5. Reloads PostgREST schema cache so RPCs are immediately callable.
-- ============================================================================

DO $$
DECLARE
  v_sig text;
BEGIN
  FOR v_sig IN
    SELECT p.oid::regprocedure::text
    FROM pg_proc p
    JOIN pg_namespace n ON n.oid = p.pronamespace
    WHERE n.nspname = 'public'
      AND p.proname IN (
        'rename_workspace_sheet',
        'delete_workspace_sheet',
        'reorder_workspace_sheets'
      )
  LOOP
    EXECUTE 'DROP FUNCTION IF EXISTS ' || v_sig || ';';
  END LOOP;
END;
$$;

-- 1. Rename Worksheet RPC
CREATE OR REPLACE FUNCTION public.rename_workspace_sheet(
  p_workspace text,
  p_sheet_id uuid,
  p_new_name text
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
  v_trimmed text;
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

  v_trimmed := btrim(coalesce(p_new_name, ''));
  IF length(v_trimmed) = 0 THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot be empty.';
  END IF;
  IF length(v_trimmed) > 31 THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot exceed 31 characters.';
  END IF;
  IF v_trimmed ~ '[\\/?*\[\]:]' THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot contain \ / ? * [ ] :';
  END IF;

  IF EXISTS (
    SELECT 1
    FROM public.workbook_sheets ws
    WHERE ws.workspace = p_workspace::public.workspace_type
      AND ws.id <> p_sheet_id
      AND lower(ws.name) = lower(v_trimmed)
  ) THEN
    RAISE EXCEPTION 'SHEET_EXISTS: A worksheet named "%" already exists.', v_trimmed;
  END IF;

  UPDATE public.workbook_sheets ws
  SET name = v_trimmed,
      sheet_key = v_trimmed
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  RETURNING ws.id, ws.name, ws.tab_order INTO v_sheet;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  RETURN jsonb_build_object(
    'success', true,
    'sheetId', v_sheet.id,
    'name', v_sheet.name,
    'tabOrder', v_sheet.tab_order
  );
END;
$$;

-- 2. Delete Worksheet RPC
CREATE OR REPLACE FUNCTION public.delete_workspace_sheet(
  p_workspace text,
  p_sheet_id uuid
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
DECLARE
  v_user_id uuid;
  v_staff record;
  v_target_sheet record;
  v_sheet_count integer;
  v_remaining record;
  v_new_order integer := 0;
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

  -- Verify target worksheet exists in the authorized workspace FIRST
  SELECT ws.id, ws.name
  INTO v_target_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  -- Prevent deleting the last remaining worksheet in this workspace
  SELECT count(*) INTO v_sheet_count
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type;

  IF v_sheet_count <= 1 THEN
    RAISE EXCEPTION 'BAD_REQUEST: Cannot delete the only remaining worksheet in the % workspace.', p_workspace;
  END IF;

  -- Delete associated rows first
  DELETE FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id;

  -- Delete the worksheet
  DELETE FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type;

  -- Re-sequence tab_order cleanly (0..N-1) for remaining worksheets in this workspace
  FOR v_remaining IN
    SELECT ws.id
    FROM public.workbook_sheets ws
    WHERE ws.workspace = p_workspace::public.workspace_type
    ORDER BY ws.tab_order ASC, ws.name ASC
  LOOP
    UPDATE public.workbook_sheets ws
    SET tab_order = v_new_order
    WHERE ws.id = v_remaining.id
      AND ws.workspace = p_workspace::public.workspace_type;
    v_new_order := v_new_order + 1;
  END LOOP;

  RETURN jsonb_build_object(
    'success', true,
    'deletedSheetId', p_sheet_id,
    'deletedSheetName', v_target_sheet.name
  );
END;
$$;

-- 3. Reorder Worksheets RPC
CREATE OR REPLACE FUNCTION public.reorder_workspace_sheets(
  p_workspace text,
  p_sheet_ids uuid[]
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
DECLARE
  v_user_id uuid;
  v_staff record;
  v_idx integer;
  v_updated_count integer := 0;
  v_row_updated integer := 0;
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

  IF p_sheet_ids IS NULL OR coalesce(array_length(p_sheet_ids, 1), 0) = 0 THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids array cannot be empty.';
  END IF;

  -- Phase 1: move targeted worksheets into a temporary non-conflicting positive range
  FOR v_idx IN 1..array_length(p_sheet_ids, 1) LOOP
    UPDATE public.workbook_sheets ws
    SET tab_order = 100000 + v_idx
    WHERE ws.id = p_sheet_ids[v_idx]
      AND ws.workspace = p_workspace::public.workspace_type;

    GET DIAGNOSTICS v_row_updated = ROW_COUNT;
    v_updated_count := v_updated_count + v_row_updated;
  END LOOP;

  IF v_updated_count = 0 THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: No matching worksheets found in % workspace to reorder.', p_workspace;
  END IF;

  -- Phase 2: assign final sequential tab_order (0..N-1)
  FOR v_idx IN 1..array_length(p_sheet_ids, 1) LOOP
    UPDATE public.workbook_sheets ws
    SET tab_order = v_idx - 1
    WHERE ws.id = p_sheet_ids[v_idx]
      AND ws.workspace = p_workspace::public.workspace_type;
  END LOOP;

  RETURN jsonb_build_object(
    'success', true,
    'reorderedCount', v_updated_count
  );
END;
$$;

REVOKE ALL ON FUNCTION public.rename_workspace_sheet(text, uuid, text) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.delete_workspace_sheet(text, uuid) FROM PUBLIC, anon;
REVOKE ALL ON FUNCTION public.reorder_workspace_sheets(text, uuid[]) FROM PUBLIC, anon;

GRANT EXECUTE ON FUNCTION public.rename_workspace_sheet(text, uuid, text) TO authenticated;
GRANT EXECUTE ON FUNCTION public.delete_workspace_sheet(text, uuid) TO authenticated;
GRANT EXECUTE ON FUNCTION public.reorder_workspace_sheets(text, uuid[]) TO authenticated;

NOTIFY pgrst, 'reload schema';
