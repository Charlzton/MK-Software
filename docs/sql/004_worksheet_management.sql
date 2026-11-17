-- ============================================================================
-- MIGRATION 004: WORKSHEET MANAGEMENT RPCS (RENAME, DELETE, REORDER)
-- ============================================================================
-- Safe, rerunnable migration adding server-side RPCs for:
--   1. public.rename_workspace_sheet
--   2. public.delete_workspace_sheet
--   3. public.reorder_workspace_sheets
-- ============================================================================

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
  -- Authenticate active staff
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

  v_trimmed := trim(both ' ' from p_new_name);
  IF v_trimmed IS NULL OR length(v_trimmed) = 0 THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot be empty.';
  END IF;
  IF length(v_trimmed) > 31 THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot exceed 31 characters.';
  END IF;
  IF v_trimmed ~ '[\\/?*\[\]:]' THEN
    RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name cannot contain \ / ? * [ ] :';
  END IF;

  -- Check for existing duplicate sheet name in this workspace
  IF EXISTS (
    SELECT 1 FROM public.workbook_sheets ws
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
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found.';
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
  v_sheet_count integer;
  v_deleted_name text;
BEGIN
  -- Authenticate active staff
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

  -- Verify at least 2 sheets exist in this workspace so we do not delete the last remaining sheet
  SELECT count(*) INTO v_sheet_count
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type;

  IF v_sheet_count <= 1 THEN
    RAISE EXCEPTION 'BAD_REQUEST: Cannot delete the only remaining worksheet in the % workspace.', p_workspace;
  END IF;

  -- Delete associated rows first
  DELETE FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id;

  -- Delete worksheet
  DELETE FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type
  RETURNING ws.name INTO v_deleted_name;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found.';
  END IF;

  RETURN jsonb_build_object(
    'success', true,
    'deletedSheetId', p_sheet_id,
    'deletedSheetName', v_deleted_name
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
BEGIN
  -- Authenticate active staff
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

  IF p_sheet_ids IS NULL OR array_length(p_sheet_ids, 1) = 0 THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids array cannot be empty.';
  END IF;

  FOR v_idx IN 1..array_length(p_sheet_ids, 1) LOOP
    UPDATE public.workbook_sheets ws
    SET tab_order = v_idx - 1
    WHERE ws.id = p_sheet_ids[v_idx]
      AND ws.workspace = p_workspace::public.workspace_type;
  END LOOP;

  RETURN jsonb_build_object(
    'success', true,
    'reorderedCount', array_length(p_sheet_ids, 1)
  );
END;
$$;

GRANT EXECUTE ON FUNCTION public.rename_workspace_sheet(text, uuid, text) TO authenticated;
GRANT EXECUTE ON FUNCTION public.delete_workspace_sheet(text, uuid) TO authenticated;
GRANT EXECUTE ON FUNCTION public.reorder_workspace_sheets(text, uuid[]) TO authenticated;
