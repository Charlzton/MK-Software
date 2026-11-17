-- ============================================================================
-- SECTION 1: READ-ONLY SCHEMA & FUNCTION INSPECTION QUERIES
-- ============================================================================
-- Run these read-only queries first in Supabase SQL Editor if you want to
-- inspect existing tables and functions without modifying anything.
--
--   SELECT table_name, column_name, data_type, udt_name, is_nullable, column_default
--   FROM information_schema.columns
--   WHERE table_schema = 'public'
--     AND table_name IN ('migration_runs', 'workbook_sheets', 'workbook_rows', 'staff_profiles', 'audit_logs')
--   ORDER BY table_name, ordinal_position;
--
--   SELECT p.proname, pg_get_functiondef(p.oid)
--   FROM pg_proc p
--   JOIN pg_namespace n ON n.oid = p.pronamespace
--   WHERE n.nspname = 'public'
--     AND p.proname IN (
--       'get_workspace_workbook',
--       'create_workspace_sheet',
--       'patch_workspace_cell',
--       'patch_workspace_cells_batch',
--       'patch_workspace_header'
--     );
--
-- ============================================================================
-- SECTION 2: EXECUTABLE MIGRATION — INITIAL WORKBOOK DATA IMPORT RPC
-- ============================================================================
-- Exact payload contract from lib/importWorkbook.ts:
--   p_payload = {
--     "reception": [{ "name": string, "sheetKey": string, "tabOrder": number, "headers": string[], "rows": (string|number)[][] }, ...],
--     "doctor":    [{ "name": string, "sheetKey": string, "tabOrder": number, "headers": string[], "rows": (string|number)[][] }, ...]
--   }
-- Schema alignment:
--   - public.migration_runs: import_batch_id, checksum, status,
--       reception_sheet_count, doctor_sheet_count, imported_row_count,
--       executed_at, completed_at, created_at
--   - public.workbook_sheets: id, workspace (public.workspace_type), sheet_key,
--       name, tab_order, headers (jsonb), header_version (bigint), import_batch_id
--   - public.workbook_rows: id, sheet_id, row_index, cells (jsonb),
--       row_version (bigint), import_batch_id, cell_versions (jsonb)
--     (workbook_rows has NO workspace column)
-- ============================================================================

CREATE OR REPLACE FUNCTION public.import_initial_workbooks(
  p_batch_key text,
  p_payload jsonb
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = pg_catalog, public, auth
AS $$
DECLARE
  v_user_id uuid;
  v_staff record;
  v_batch_id uuid;
  v_batch_record record;
  v_sheet_record record;

  v_existing_sheet_count integer := 0;
  v_existing_row_count integer := 0;
  v_batch_sheet_count integer := 0;
  v_batch_row_count integer := 0;

  v_expected_reception_sheets integer := 0;
  v_expected_reception_rows integer := 0;
  v_expected_doctor_sheets integer := 0;
  v_expected_doctor_rows integer := 0;

  v_reception_sheets integer := 0;
  v_reception_rows integer := 0;
  v_doctor_sheets integer := 0;
  v_doctor_rows integer := 0;

  v_seen_sheet_keys text[];
  v_seen_tab_orders integer[];

  v_matching_sheet_count integer := 0;
  v_expected_sheet_row_count integer := 0;
  v_actual_sheet_row_count integer := 0;
  v_matched_row_count integer := 0;

  v_workspace_key text;
  v_sheet_elem jsonb;
  v_sheet_id uuid;
  v_sheet_name text;
  v_sheet_key text;
  v_tab_order integer;
  v_headers jsonb;
  v_rows jsonb;
  v_row_elem jsonb;
  v_row_idx integer;
BEGIN
  -- 1. Require authenticated active ADMIN staff user
  v_user_id := auth.uid();
  IF v_user_id IS NULL THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Authentication required.';
  END IF;

  SELECT user_id, role::text AS role_text, display_name, is_active
  INTO v_staff
  FROM public.staff_profiles
  WHERE user_id = v_user_id;

  IF NOT FOUND OR v_staff.is_active IS NOT TRUE THEN
    RAISE EXCEPTION 'UNAUTHORIZED: Inactive or missing staff profile.';
  END IF;

  IF v_staff.role_text <> 'admin' THEN
    RAISE EXCEPTION 'FORBIDDEN_ROLE: Only admin staff may initiate the initial workbook import.';
  END IF;

  -- 2. Validate batch key and payload structure strictly (separate array checks before length/elements)
  IF p_batch_key IS NULL OR btrim(p_batch_key) = '' THEN
    RAISE EXCEPTION 'INVALID_BATCH_KEY: p_batch_key is required.';
  END IF;

  IF p_payload IS NULL OR jsonb_typeof(p_payload) IS DISTINCT FROM 'object' THEN
    RAISE EXCEPTION 'INVALID_PAYLOAD: p_payload must be a JSON object.';
  END IF;

  IF jsonb_typeof(p_payload->'reception') IS DISTINCT FROM 'array' THEN
    RAISE EXCEPTION 'INVALID_PAYLOAD: p_payload.reception must be a JSON array.';
  END IF;

  IF jsonb_typeof(p_payload->'doctor') IS DISTINCT FROM 'array' THEN
    RAISE EXCEPTION 'INVALID_PAYLOAD: p_payload.doctor must be a JSON array.';
  END IF;

  v_expected_reception_sheets := jsonb_array_length(p_payload->'reception');
  IF v_expected_reception_sheets = 0 THEN
    RAISE EXCEPTION 'INVALID_PAYLOAD: p_payload.reception array cannot be empty.';
  END IF;

  v_expected_doctor_sheets := jsonb_array_length(p_payload->'doctor');
  IF v_expected_doctor_sheets = 0 THEN
    RAISE EXCEPTION 'INVALID_PAYLOAD: p_payload.doctor array cannot be empty.';
  END IF;

  -- Pre-validate every worksheet object, row types, nonnegative integer tabOrder, and per-workspace uniqueness
  FOREACH v_workspace_key IN ARRAY ARRAY['reception', 'doctor']
  LOOP
    v_seen_sheet_keys := ARRAY[]::text[];
    v_seen_tab_orders := ARRAY[]::integer[];

    FOR v_sheet_elem IN
      SELECT value
      FROM jsonb_array_elements(p_payload->v_workspace_key)
    LOOP
      IF jsonb_typeof(v_sheet_elem) IS DISTINCT FROM 'object' THEN
        RAISE EXCEPTION 'INVALID_SHEET: Every worksheet entry in % must be a JSON object.', v_workspace_key;
      END IF;

      v_sheet_name := v_sheet_elem->>'name';
      IF v_sheet_name IS NULL OR btrim(v_sheet_name) = '' THEN
        RAISE EXCEPTION 'INVALID_SHEET_NAME: Worksheet name is required in %.', v_workspace_key;
      END IF;

      v_sheet_key := v_sheet_elem->>'sheetKey';
      IF v_sheet_key IS NULL OR btrim(v_sheet_key) = '' THEN
        RAISE EXCEPTION 'INVALID_SHEET_KEY: Worksheet sheetKey is required for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      IF v_sheet_key = ANY(v_seen_sheet_keys) THEN
        RAISE EXCEPTION 'INVALID_SHEET_KEY: Duplicate sheetKey "%" in %.', v_sheet_key, v_workspace_key;
      END IF;
      v_seen_sheet_keys := array_append(v_seen_sheet_keys, v_sheet_key);

      IF jsonb_typeof(v_sheet_elem->'tabOrder') IS DISTINCT FROM 'number' THEN
        RAISE EXCEPTION 'INVALID_TAB_ORDER: Numeric tabOrder is required for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      IF (v_sheet_elem->>'tabOrder') !~ '^[0-9]+$' THEN
        RAISE EXCEPTION 'INVALID_TAB_ORDER: tabOrder must be a nonnegative integer for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      v_tab_order := (v_sheet_elem->>'tabOrder')::integer;
      IF v_tab_order < 0 THEN
        RAISE EXCEPTION 'INVALID_TAB_ORDER: tabOrder must be >= 0 for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      IF v_tab_order = ANY(v_seen_tab_orders) THEN
        RAISE EXCEPTION 'INVALID_TAB_ORDER: Duplicate tabOrder % in %.', v_tab_order, v_workspace_key;
      END IF;
      v_seen_tab_orders := array_append(v_seen_tab_orders, v_tab_order);

      IF jsonb_typeof(v_sheet_elem->'headers') IS DISTINCT FROM 'array' THEN
        RAISE EXCEPTION 'INVALID_HEADERS: headers array is required for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      IF EXISTS (
        SELECT 1
        FROM jsonb_array_elements(v_sheet_elem->'headers') AS h(val)
        WHERE jsonb_typeof(h.val) IS DISTINCT FROM 'string'
      ) THEN
        RAISE EXCEPTION 'INVALID_HEADERS: Every header in "%" (%) must be a JSON string.', v_sheet_name, v_workspace_key;
      END IF;

      IF jsonb_typeof(v_sheet_elem->'rows') IS DISTINCT FROM 'array' THEN
        RAISE EXCEPTION 'INVALID_ROWS: rows array is required for "%" in %.', v_sheet_name, v_workspace_key;
      END IF;

      IF EXISTS (
        SELECT 1
        FROM jsonb_array_elements(v_sheet_elem->'rows') AS r(val)
        WHERE jsonb_typeof(r.val) IS DISTINCT FROM 'array'
      ) THEN
        RAISE EXCEPTION 'INVALID_ROW_DATA: Every row in "%" (%) must be a JSON array.', v_sheet_name, v_workspace_key;
      END IF;

      IF EXISTS (
        SELECT 1
        FROM jsonb_array_elements(v_sheet_elem->'rows') AS r(val),
             jsonb_array_elements(r.val) AS c(val)
        WHERE jsonb_typeof(c.val) IS NULL
           OR jsonb_typeof(c.val) NOT IN ('string', 'number')
      ) THEN
        RAISE EXCEPTION 'INVALID_CELL_DATA: Every cell in "%" (%) must be a JSON string or JSON number.', v_sheet_name, v_workspace_key;
      END IF;

      IF v_workspace_key = 'reception' THEN
        v_expected_reception_rows := v_expected_reception_rows + jsonb_array_length(v_sheet_elem->'rows');
      ELSE
        v_expected_doctor_rows := v_expected_doctor_rows + jsonb_array_length(v_sheet_elem->'rows');
      END IF;
    END LOOP;
  END LOOP;

  -- 3. Serialize concurrent import attempts inside this transaction
  PERFORM pg_advisory_xact_lock(hashtext('mk_eye_center_initial_workbook_import'));

  -- 4. Inspect existing live workbook data across BOTH workspaces (joining workbook_rows via sheet_id)
  SELECT COUNT(*)
  INTO v_existing_sheet_count
  FROM public.workbook_sheets ws
  WHERE ws.workspace IN ('reception'::public.workspace_type, 'doctor'::public.workspace_type);

  SELECT COUNT(*)
  INTO v_existing_row_count
  FROM public.workbook_rows wr
  JOIN public.workbook_sheets ws ON ws.id = wr.sheet_id
  WHERE ws.workspace IN ('reception'::public.workspace_type, 'doctor'::public.workspace_type);

  -- 5. Strict idempotency check via migration_runs.checksum, import_batch_id, counts, AND full content verification
  SELECT
    mr.import_batch_id,
    mr.reception_sheet_count,
    mr.doctor_sheet_count,
    mr.imported_row_count
  INTO v_batch_record
  FROM public.migration_runs mr
  WHERE mr.checksum = p_batch_key
    AND mr.status = 'completed'
  ORDER BY mr.completed_at DESC NULLS LAST, mr.created_at DESC NULLS LAST
  LIMIT 1;

  IF FOUND AND v_batch_record.import_batch_id IS NOT NULL THEN
    v_batch_id := v_batch_record.import_batch_id;

    SELECT
      COUNT(*) FILTER (WHERE ws.workspace = 'reception'::public.workspace_type),
      COUNT(*) FILTER (WHERE ws.workspace = 'doctor'::public.workspace_type)
    INTO v_reception_sheets, v_doctor_sheets
    FROM public.workbook_sheets ws
    WHERE ws.import_batch_id = v_batch_id
      AND ws.workspace IN ('reception'::public.workspace_type, 'doctor'::public.workspace_type);

    SELECT
      COUNT(*) FILTER (WHERE ws.workspace = 'reception'::public.workspace_type),
      COUNT(*) FILTER (WHERE ws.workspace = 'doctor'::public.workspace_type)
    INTO v_reception_rows, v_doctor_rows
    FROM public.workbook_rows wr
    JOIN public.workbook_sheets ws ON ws.id = wr.sheet_id
    WHERE wr.import_batch_id = v_batch_id
      AND ws.import_batch_id = v_batch_id
      AND ws.workspace IN ('reception'::public.workspace_type, 'doctor'::public.workspace_type);

    v_batch_sheet_count := v_reception_sheets + v_doctor_sheets;
    v_batch_row_count := v_reception_rows + v_doctor_rows;

    IF v_batch_record.reception_sheet_count <> v_expected_reception_sheets
       OR v_batch_record.doctor_sheet_count <> v_expected_doctor_sheets
       OR v_batch_record.imported_row_count <> (v_expected_reception_rows + v_expected_doctor_rows)
       OR v_reception_sheets <> v_expected_reception_sheets
       OR v_doctor_sheets <> v_expected_doctor_sheets
       OR v_reception_rows <> v_expected_reception_rows
       OR v_doctor_rows <> v_expected_doctor_rows
       OR v_existing_sheet_count <> v_batch_sheet_count
       OR v_existing_row_count <> v_batch_row_count THEN
      RAISE EXCEPTION 'IMPORT_CONFLICT: Migration batch "%" exists but live counts (% sheets, % rows) do not match expected (% sheets, % rows).',
        p_batch_key,
        v_existing_sheet_count,
        v_existing_row_count,
        v_expected_reception_sheets + v_expected_doctor_sheets,
        v_expected_reception_rows + v_expected_doctor_rows;
    END IF;

    -- Verify every worksheet (including 0-row worksheets) and every row's row_index & cells against the supplied payload
    FOREACH v_workspace_key IN ARRAY ARRAY['reception', 'doctor']
    LOOP
      FOR v_sheet_elem IN
        SELECT value
        FROM jsonb_array_elements(p_payload->v_workspace_key)
      LOOP
        v_sheet_name := v_sheet_elem->>'name';
        v_sheet_key := v_sheet_elem->>'sheetKey';
        v_tab_order := (v_sheet_elem->>'tabOrder')::integer;
        v_headers := v_sheet_elem->'headers';
        v_rows := v_sheet_elem->'rows';
        v_expected_sheet_row_count := jsonb_array_length(v_rows);

        SELECT COUNT(*)
        INTO v_matching_sheet_count
        FROM public.workbook_sheets ws
        WHERE ws.import_batch_id = v_batch_id
          AND ws.workspace = v_workspace_key::public.workspace_type
          AND ws.sheet_key = v_sheet_key
          AND ws.tab_order = v_tab_order;

        IF v_matching_sheet_count <> 1 THEN
          RAISE EXCEPTION 'IMPORT_CONFLICT: Worksheet "%" (sheet_key=%, tab_order=%) in % is missing or duplicated in batch %.',
            v_sheet_name, v_sheet_key, v_tab_order, v_workspace_key, v_batch_id;
        END IF;

        SELECT ws.id, ws.name, ws.headers
        INTO v_sheet_record
        FROM public.workbook_sheets ws
        WHERE ws.import_batch_id = v_batch_id
          AND ws.workspace = v_workspace_key::public.workspace_type
          AND ws.sheet_key = v_sheet_key
          AND ws.tab_order = v_tab_order;

        IF v_sheet_record.name IS DISTINCT FROM v_sheet_name
           OR v_sheet_record.headers IS DISTINCT FROM v_headers THEN
          RAISE EXCEPTION 'IMPORT_CONFLICT: Worksheet "%" (%) name or headers differ from batch %.',
            v_sheet_name, v_workspace_key, v_batch_id;
        END IF;

        SELECT COUNT(*)
        INTO v_actual_sheet_row_count
        FROM public.workbook_rows wr
        JOIN public.workbook_sheets ws ON ws.id = wr.sheet_id
        WHERE ws.id = v_sheet_record.id
          AND ws.workspace = v_workspace_key::public.workspace_type
          AND wr.import_batch_id = v_batch_id;

        IF v_actual_sheet_row_count <> v_expected_sheet_row_count THEN
          RAISE EXCEPTION 'IMPORT_CONFLICT: Worksheet "%" (%) row count % differs from expected % in batch %.',
            v_sheet_name, v_workspace_key, v_actual_sheet_row_count, v_expected_sheet_row_count, v_batch_id;
        END IF;

        IF v_expected_sheet_row_count > 0 THEN
          SELECT COUNT(DISTINCT wr.row_index)
          INTO v_matched_row_count
          FROM jsonb_array_elements(v_rows) WITH ORDINALITY AS exp(cells, ord)
          JOIN public.workbook_rows wr
            ON wr.sheet_id = v_sheet_record.id
           AND wr.import_batch_id = v_batch_id
           AND wr.row_index = (exp.ord - 1)::integer
           AND wr.cells = exp.cells
          JOIN public.workbook_sheets ws
            ON ws.id = wr.sheet_id
           AND ws.workspace = v_workspace_key::public.workspace_type;

          IF v_matched_row_count <> v_expected_sheet_row_count THEN
            RAISE EXCEPTION 'IMPORT_CONFLICT: Row index or cell contents in worksheet "%" (%) do not match batch %.',
              v_sheet_name, v_workspace_key, v_batch_id;
          END IF;
        END IF;
      END LOOP;
    END LOOP;

    RETURN jsonb_build_object(
      'status', 'already_imported',
      'already_imported', true,
      'import_batch_id', v_batch_id,
      'reception_sheets', v_reception_sheets,
      'reception_rows', v_reception_rows,
      'doctor_sheets', v_doctor_sheets,
      'doctor_rows', v_doctor_rows,
      'total_sheets', v_batch_sheet_count,
      'total_rows', v_batch_row_count
    );
  END IF;

  -- 6. Refuse import if EITHER workspace already contains live sheets or rows
  IF v_existing_sheet_count > 0 OR v_existing_row_count > 0 THEN
    RAISE EXCEPTION 'IMPORT_CONFLICT: Existing workbook data already exists (% sheets, % rows). Initial import refused to protect live data.',
      v_existing_sheet_count, v_existing_row_count;
  END IF;

  -- 7. Record migration run (any subsequent error rolls back this insert atomically)
  INSERT INTO public.migration_runs (
    checksum,
    status,
    reception_sheet_count,
    doctor_sheet_count,
    imported_row_count,
    executed_at
  )
  VALUES (
    p_batch_key,
    'started',
    0,
    0,
    0,
    now()
  )
  RETURNING import_batch_id INTO v_batch_id;

  -- Reset counters before insertion loop
  v_reception_sheets := 0;
  v_reception_rows := 0;
  v_doctor_sheets := 0;
  v_doctor_rows := 0;

  -- 8. Import Reception and Doctor workspaces atomically (preserving 0-row worksheets, exact keys, tabOrder, headers, rows)
  FOREACH v_workspace_key IN ARRAY ARRAY['reception', 'doctor']
  LOOP
    FOR v_sheet_elem IN
      SELECT value
      FROM jsonb_array_elements(p_payload->v_workspace_key)
    LOOP
      v_sheet_name := v_sheet_elem->>'name';
      v_sheet_key := v_sheet_elem->>'sheetKey';
      v_tab_order := (v_sheet_elem->>'tabOrder')::integer;
      v_headers := v_sheet_elem->'headers';
      v_rows := v_sheet_elem->'rows';

      INSERT INTO public.workbook_sheets (
        workspace,
        sheet_key,
        name,
        tab_order,
        headers,
        header_version,
        import_batch_id
      )
      VALUES (
        v_workspace_key::public.workspace_type,
        v_sheet_key,
        v_sheet_name,
        v_tab_order,
        v_headers,
        1,
        v_batch_id
      )
      RETURNING id INTO v_sheet_id;

      IF v_workspace_key = 'reception' THEN
        v_reception_sheets := v_reception_sheets + 1;
      ELSE
        v_doctor_sheets := v_doctor_sheets + 1;
      END IF;

      -- Insert rows preserving exact row_index (0..N-1), blanks, and ragged lengths;
      -- 0-row worksheets execute 0 iterations while keeping their workbook_sheets record.
      v_row_idx := 0;
      FOR v_row_elem IN
        SELECT value
        FROM jsonb_array_elements(v_rows)
      LOOP
        IF jsonb_typeof(v_row_elem) IS DISTINCT FROM 'array' THEN
          RAISE EXCEPTION 'INVALID_ROW_DATA: Row % in "%" (%) must be a JSON array.', v_row_idx, v_sheet_name, v_workspace_key;
        END IF;

        INSERT INTO public.workbook_rows (
          sheet_id,
          row_index,
          cells,
          row_version,
          import_batch_id,
          cell_versions
        )
        VALUES (
          v_sheet_id,
          v_row_idx,
          v_row_elem,
          1,
          v_batch_id,
          '{}'::jsonb
        );

        IF v_workspace_key = 'reception' THEN
          v_reception_rows := v_reception_rows + 1;
        ELSE
          v_doctor_rows := v_doctor_rows + 1;
        END IF;

        v_row_idx := v_row_idx + 1;
      END LOOP;
    END LOOP;
  END LOOP;

  -- 9. Mark migration run completed using import_batch_id and actual inserted counts
  UPDATE public.migration_runs
  SET
    status = 'completed',
    reception_sheet_count = v_reception_sheets,
    doctor_sheet_count = v_doctor_sheets,
    imported_row_count = v_reception_rows + v_doctor_rows,
    completed_at = now()
  WHERE import_batch_id = v_batch_id;

  RETURN jsonb_build_object(
    'status', 'imported',
    'already_imported', false,
    'import_batch_id', v_batch_id,
    'reception_sheets', v_reception_sheets,
    'reception_rows', v_reception_rows,
    'doctor_sheets', v_doctor_sheets,
    'doctor_rows', v_doctor_rows,
    'total_sheets', v_reception_sheets + v_doctor_sheets,
    'total_rows', v_reception_rows + v_doctor_rows
  );
END;
$$;

-- Restrict execution strictly to authenticated users (function body enforces role = 'admin')
REVOKE ALL ON FUNCTION public.import_initial_workbooks(text, jsonb) FROM PUBLIC, anon;
GRANT EXECUTE ON FUNCTION public.import_initial_workbooks(text, jsonb) TO authenticated;
