-- ============================================================================
-- MIGRATION 005: HARDENED WORKSHEET MANAGEMENT RPCS (DELETE, REORDER, RENAME)
-- ============================================================================
-- Safe, rerunnable migration fixing worksheet deletion and drag-and-drop
-- reordering in Supabase for both Reception and Doctor workspaces.
--
-- What this fixes:
--   1. Drops only the exact existing RPC signatures (no broad pg_proc sweep).
--   2. rename_workspace_sheet:
--      - Locks all worksheets in the workspace deterministically to prevent
--        concurrent rename races to the same name.
--      - Hardened duplicate check, validation, and permission enforcement.
--   3. delete_workspace_sheet:
--      - Authenticates active staff and verifies workspace authorization.
--      - Locks all worksheets in the workspace deterministically BEFORE
--        checking count > 1 (prevents concurrent deletion of the last sheet).
--      - Locks associated workbook_rows in deterministic order before delete.
--      - Safely deletes associated workbook_rows, deletes the workbook_sheets
--        record, and re-sequences remaining tab_order values (0..N-1) via a
--        collision-safe two-phase update.
--   4. reorder_workspace_sheets:
--      - Authenticates active staff and verifies workspace authorization.
--      - Validates p_sheet_ids has no NULLs, no duplicates, and matches every
--        worksheet in the requested workspace 1-to-1 under FOR UPDATE locks.
--      - Atomically updates tab_order using a collision-safe dynamic two-phase
--        offset so unique (workspace, tab_order) constraints never fail.
--   5. Reloads PostgREST schema cache so RPCs are immediately callable.
-- ============================================================================

DROP FUNCTION IF EXISTS public.rename_workspace_sheet(text, uuid, text);
DROP FUNCTION IF EXISTS public.delete_workspace_sheet(text, uuid);
DROP FUNCTION IF EXISTS public.reorder_workspace_sheets(text, uuid[]);

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

  -- Lock workspace worksheets in deterministic order to serialize concurrent renames/deletes
  PERFORM ws.id
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type
  ORDER BY ws.tab_order, ws.id
  FOR UPDATE;

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
  v_max_tab_order integer;
  v_temp_offset integer;
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

  -- Lock ALL worksheets in this workspace deterministically BEFORE counting or deleting
  -- so concurrent delete/reorder calls cannot race past the last-sheet guard.
  PERFORM ws.id
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type
  ORDER BY ws.tab_order, ws.id
  FOR UPDATE;

  -- Verify target worksheet exists in the authorized workspace
  SELECT ws.id, ws.name
  INTO v_target_sheet
  FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type;

  IF NOT FOUND THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: Worksheet not found in % workspace.', p_workspace;
  END IF;

  -- Prevent deleting the last remaining worksheet in this workspace
  SELECT count(*)::integer, COALESCE(MAX(ws.tab_order), 0)
  INTO v_sheet_count, v_max_tab_order
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type;

  IF v_sheet_count <= 1 THEN
    RAISE EXCEPTION 'BAD_REQUEST: Cannot delete the only remaining worksheet in the % workspace.', p_workspace;
  END IF;

  -- Lock associated rows in canonical order before deleting
  PERFORM wr.id
  FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id
  ORDER BY wr.row_index, wr.id
  FOR UPDATE;

  -- Delete associated rows first
  DELETE FROM public.workbook_rows wr
  WHERE wr.sheet_id = p_sheet_id;

  -- Delete the worksheet
  DELETE FROM public.workbook_sheets ws
  WHERE ws.id = p_sheet_id
    AND ws.workspace = p_workspace::public.workspace_type;

  -- Collision-safe two-phase re-sequence (0..N-1) for remaining worksheets in this workspace
  v_temp_offset := GREATEST(1000000, v_max_tab_order + 1000);

  WITH ordered AS (
    SELECT
      ws.id,
      (row_number() OVER (ORDER BY ws.tab_order ASC, ws.name ASC, ws.id ASC))::integer AS seq
    FROM public.workbook_sheets ws
    WHERE ws.workspace = p_workspace::public.workspace_type
  )
  UPDATE public.workbook_sheets ws
  SET tab_order = v_temp_offset + ordered.seq
  FROM ordered
  WHERE ws.id = ordered.id
    AND ws.workspace = p_workspace::public.workspace_type;

  WITH ordered AS (
    SELECT
      ws.id,
      (row_number() OVER (ORDER BY ws.tab_order ASC, ws.name ASC, ws.id ASC))::integer AS seq
    FROM public.workbook_sheets ws
    WHERE ws.workspace = p_workspace::public.workspace_type
  )
  UPDATE public.workbook_sheets ws
  SET tab_order = ordered.seq - 1
  FROM ordered
  WHERE ws.id = ordered.id
    AND ws.workspace = p_workspace::public.workspace_type;

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
  v_input_count integer;
  v_distinct_count integer;
  v_workspace_sheet_count integer;
  v_matched_count integer;
  v_max_tab_order integer;
  v_temp_offset integer;
  v_updated_count integer := 0;
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

  IF p_sheet_ids IS NULL OR coalesce(array_ndims(p_sheet_ids), 0) <> 1 THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids must be a non-empty 1-dimensional UUID array.';
  END IF;

  v_input_count := coalesce(array_length(p_sheet_ids, 1), 0);
  IF v_input_count = 0 THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids array cannot be empty.';
  END IF;

  IF EXISTS (
    SELECT 1
    FROM unnest(p_sheet_ids) AS u(id)
    WHERE u.id IS NULL
  ) THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids cannot contain NULL entries.';
  END IF;

  SELECT COUNT(DISTINCT u.id)::integer
  INTO v_distinct_count
  FROM unnest(p_sheet_ids) AS u(id);

  IF v_distinct_count <> v_input_count THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids contains duplicate worksheet IDs.';
  END IF;

  -- Lock ALL worksheets in the workspace deterministically BEFORE validating completeness
  PERFORM ws.id
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type
  ORDER BY ws.tab_order, ws.id
  FOR UPDATE;

  SELECT COUNT(*)::integer, COALESCE(MAX(ws.tab_order), 0)
  INTO v_workspace_sheet_count, v_max_tab_order
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type;

  IF v_workspace_sheet_count = 0 THEN
    RAISE EXCEPTION 'SHEET_NOT_FOUND: No worksheets found in % workspace.', p_workspace;
  END IF;

  SELECT COUNT(*)::integer
  INTO v_matched_count
  FROM public.workbook_sheets ws
  WHERE ws.workspace = p_workspace::public.workspace_type
    AND ws.id = ANY(p_sheet_ids);

  IF v_matched_count <> v_input_count OR v_matched_count <> v_workspace_sheet_count THEN
    RAISE EXCEPTION 'BAD_REQUEST: p_sheet_ids must contain the exact set of % worksheets in % workspace (matched %, supplied %).',
      v_workspace_sheet_count, p_workspace, v_matched_count, v_input_count;
  END IF;

  -- Collision-safe two-phase update across all workspace sheets
  v_temp_offset := GREATEST(1000000, v_max_tab_order + 1000);

  -- Phase 1: move all worksheets into a disjoint temporary positive range
  UPDATE public.workbook_sheets ws
  SET tab_order = v_temp_offset + u.ord::integer
  FROM unnest(p_sheet_ids) WITH ORDINALITY AS u(id, ord)
  WHERE ws.id = u.id
    AND ws.workspace = p_workspace::public.workspace_type;

  -- Phase 2: assign final sequential tab_order (0..N-1)
  UPDATE public.workbook_sheets ws
  SET tab_order = (u.ord - 1)::integer
  FROM unnest(p_sheet_ids) WITH ORDINALITY AS u(id, ord)
  WHERE ws.id = u.id
    AND ws.workspace = p_workspace::public.workspace_type;

  GET DIAGNOSTICS v_updated_count = ROW_COUNT;

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
