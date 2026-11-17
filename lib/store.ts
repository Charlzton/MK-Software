import type { SupabaseClient } from '@supabase/supabase-js';
import type { UserRole } from '@/lib/types';

export type WorkspaceType = 'reception' | 'doctor';

export interface WorkspaceSheetState {
  id: string;
  name: string;
  sheetKey?: string;
  tabOrder: number;
  headers: string[];
  headerVersion: number;
  rows: (string | number)[][];
  rowVersions: Record<number, number>;
  cellVersions: Record<number, Record<number, number>>;
  rowUpdatedAt?: Record<number, string>;
}

export interface WorkspaceWorkbookResult {
  workspace: WorkspaceType;
  sheets: Record<string, WorkspaceSheetState>;
  sheetNames: string[];
  lastSavedAt: string | null;
}

export interface CellEditInput {
  rowIndex: number;
  colIndex: number;
  value: string | number;
  expectedCellVersion: number;
}

export interface PatchCellResult {
  sheetId: string;
  rowIndex: number;
  colIndex: number;
  newCellVersion: number;
  newRowVersion?: number;
  updatedAt?: string;
}

export interface PatchCellsBatchResult {
  sheetId: string;
  editCount: number;
  updatedCells: Array<{
    rowIndex: number;
    colIndex: number;
    newCellVersion: number;
    newRowVersion?: number;
  }>;
  updatedAt?: string;
}

export interface PatchHeaderResult {
  sheetId: string;
  colIndex: number;
  headers: string[];
  newHeaderVersion: number;
  updatedAt?: string;
}

export class StoreRpcError extends Error {
  status: number;
  code: string;
  details?: string;

  constructor(status: number, code: string, message: string, details?: string) {
    super(message);
    this.name = 'StoreRpcError';
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

export function canRoleAccessWorkspace(role: UserRole, workspace: WorkspaceType): boolean {
  if (role === 'admin') return true;
  if (role === 'reception') return workspace === 'reception';
  if (role === 'doctor') return workspace === 'doctor';
  return false;
}

export function validateWorksheetName(rawName: string): { valid: true; name: string } | { valid: false; error: string } {
  const trimmed = rawName.trim();
  if (!trimmed) {
    return { valid: false, error: 'Worksheet name cannot be empty.' };
  }
  if (/[\\/?*[\]:]/.test(trimmed)) {
    return {
      valid: false,
      error: 'Worksheet name cannot contain any of the following characters: \\ / ? * [ ] :'
    };
  }
  if (trimmed.length > 31) {
    return { valid: false, error: 'Worksheet name cannot exceed 31 characters.' };
  }
  return { valid: true, name: trimmed };
}

function mapRpcError(error: { code?: string; message?: string; details?: string; hint?: string }): StoreRpcError {
  const rawMessage = error.message || 'Database RPC error.';
  const rawCode = error.code || '';

  if (
    rawCode === 'PGRST202' ||
    /Could not find the function public\.import_initial_workbooks/i.test(rawMessage)
  ) {
    return new StoreRpcError(
      501,
      'MIGRATION_REQUIRED',
      'Database RPC import_initial_workbooks is not installed yet. Run the SQL migration docs/sql/002_initial_workbook_import.sql in Supabase SQL Editor first.',
      rawMessage
    );
  }
  if (/IMPORT_CONFLICT/i.test(rawMessage)) {
    return new StoreRpcError(409, 'IMPORT_CONFLICT', rawMessage, rawMessage);
  }
  if (/INVALID_IMPORT_BATCH|row import batch/i.test(rawMessage)) {
    return new StoreRpcError(
      409,
      'IMPORT_BATCH_MISMATCH',
      'Database save error: row import batch validation mismatch. Run docs/sql/003_fix_patch_workspace_cell_import_batch.sql in Supabase SQL Editor and retry.',
      rawMessage
    );
  }
  if (/CONFLICT|VERSION_CONFLICT|stale/i.test(rawMessage)) {
    return new StoreRpcError(
      409,
      'VERSION_CONFLICT',
      'Version conflict: another user modified this data since you loaded it. Reload the latest data and retry your edit.',
      rawMessage
    );
  }
  if (/FORBIDDEN|FORBIDDEN_ROLE|not authorized/i.test(rawMessage) || rawCode === '42501' || /permission denied/i.test(rawMessage)) {
    return new StoreRpcError(403, 'FORBIDDEN', rawMessage);
  }
  if (/UNAUTHORIZED|Authentication required|inactive staff/i.test(rawMessage)) {
    return new StoreRpcError(401, 'UNAUTHORIZED', rawMessage);
  }
  if (/SHEET_NOT_FOUND|not found/i.test(rawMessage)) {
    return new StoreRpcError(404, 'NOT_FOUND', rawMessage);
  }
  if (/CONCURRENCY_REQUIRED|INVALID_|SHEET_EXISTS|already exists/i.test(rawMessage)) {
    return new StoreRpcError(400, 'BAD_REQUEST', rawMessage);
  }

  return new StoreRpcError(500, rawCode || 'RPC_ERROR', rawMessage, error.details || error.hint);
}

function normalizeCellScalar(val: unknown): string | number {
  if (typeof val === 'number' && Number.isFinite(val)) return val;
  if (typeof val === 'string') return val;
  if (val === null || val === undefined) return '';
  return String(val);
}

function normalizeCellVersionsMap(raw: unknown): Record<number, number> {
  const result: Record<number, number> = {};
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return result;

  for (const [key, val] of Object.entries(raw as Record<string, unknown>)) {
    const colIdx = Number(key);
    const ver = Number(val);
    if (Number.isInteger(colIdx) && colIdx >= 0 && Number.isFinite(ver) && ver >= 0) {
      result[colIdx] = ver;
    }
  }
  return result;
}

interface RawRowRecord {
  row_index?: unknown;
  rowIndex?: unknown;
  cells?: unknown;
  row_version?: unknown;
  rowVersion?: unknown;
  cell_versions?: unknown;
  cellVersions?: unknown;
  updated_at?: unknown;
  updatedAt?: unknown;
}

interface RawSheetRecord {
  id?: unknown;
  sheetId?: unknown;
  name?: unknown;
  sheet_key?: unknown;
  sheetKey?: unknown;
  tab_order?: unknown;
  tabOrder?: unknown;
  headers?: unknown;
  header_version?: unknown;
  headerVersion?: unknown;
  rows?: unknown;
  updated_at?: unknown;
  updatedAt?: unknown;
}

function normalizeSheetRecord(rawSheet: RawSheetRecord, fallbackName: string, fallbackOrder: number): WorkspaceSheetState {
  const id = String(rawSheet.id ?? rawSheet.sheetId ?? '');
  const name = String(rawSheet.name ?? fallbackName);
  const sheetKeyRaw = rawSheet.sheet_key ?? rawSheet.sheetKey;
  const tabOrder = Number(rawSheet.tab_order ?? rawSheet.tabOrder ?? fallbackOrder);
  const headerVersion = Number(rawSheet.header_version ?? rawSheet.headerVersion ?? 1);
  const headers = Array.isArray(rawSheet.headers)
    ? rawSheet.headers.map((h) => (typeof h === 'string' ? h : String(h ?? '')))
    : [];

  const rows: (string | number)[][] = [];
  const rowVersions: Record<number, number> = {};
  const cellVersions: Record<number, Record<number, number>> = {};
  const rowUpdatedAt: Record<number, string> = {};

  if (Array.isArray(rawSheet.rows)) {
    const parsedRows = rawSheet.rows
      .map(
        (
          entry,
          idx
        ): {
          rowIndex: number;
          cells: (string | number)[];
          rowVersion: number;
          cellVersions: Record<number, number>;
          updatedAt?: string;
        } | null => {
          if (Array.isArray(entry)) {
            return {
              rowIndex: idx,
              cells: entry.map(normalizeCellScalar),
              rowVersion: 1,
              cellVersions: {}
            };
          }
          if (entry && typeof entry === 'object') {
            const rec = entry as RawRowRecord;
            const rowIndexVal = Number(rec.row_index ?? rec.rowIndex ?? idx);
            if (!Number.isInteger(rowIndexVal) || rowIndexVal < 0) return null;
            const cellsArr = Array.isArray(rec.cells) ? rec.cells.map(normalizeCellScalar) : [];
            const rowVer = Number(rec.row_version ?? rec.rowVersion ?? 1);
            const cellVerMap = normalizeCellVersionsMap(rec.cell_versions ?? rec.cellVersions);
            const rawUpdated = rec.updated_at ?? rec.updatedAt;
            const updatedAt =
              typeof rawUpdated === 'string' && rawUpdated.trim() !== '' ? rawUpdated : undefined;
            return {
              rowIndex: rowIndexVal,
              cells: cellsArr,
              rowVersion: Number.isFinite(rowVer) && rowVer >= 1 ? rowVer : 1,
              cellVersions: cellVerMap,
              updatedAt
            };
          }
          return null;
        }
      )
      .filter((r): r is NonNullable<typeof r> => r !== null)
      .sort((a, b) => a.rowIndex - b.rowIndex);

    for (const r of parsedRows) {
      while (rows.length < r.rowIndex) {
        rows.push([]);
      }
      rows[r.rowIndex] = r.cells;
      rowVersions[r.rowIndex] = r.rowVersion;
      cellVersions[r.rowIndex] = r.cellVersions;
      if (r.updatedAt) {
        rowUpdatedAt[r.rowIndex] = r.updatedAt;
      }
    }
  }

  return {
    id,
    name,
    ...(typeof sheetKeyRaw === 'string' ? { sheetKey: sheetKeyRaw } : {}),
    tabOrder: Number.isFinite(tabOrder) ? tabOrder : fallbackOrder,
    headers,
    headerVersion: Number.isFinite(headerVersion) && headerVersion >= 1 ? headerVersion : 1,
    rows,
    rowVersions,
    cellVersions,
    rowUpdatedAt
  };
}

export async function loadWorkspaceWorkbook(
  supabase: SupabaseClient,
  workspace: WorkspaceType
): Promise<WorkspaceWorkbookResult> {
  const { data, error } = await supabase.rpc('get_workspace_workbook', {
    p_workspace: workspace
  });

  if (error) {
    throw mapRpcError(error);
  }

  const normalizedList: WorkspaceSheetState[] = [];
  let lastSavedAt: string | null = null;

  if (data && typeof data === 'object' && !Array.isArray(data)) {
    const obj = data as {
      sheets?: unknown;
      lastSavedAt?: unknown;
      last_saved_at?: unknown;
    };
    const ts = obj.lastSavedAt ?? obj.last_saved_at;
    if (typeof ts === 'string') {
      lastSavedAt = ts;
    }

    if (Array.isArray(obj.sheets)) {
      obj.sheets.forEach((s, idx) => {
        if (s && typeof s === 'object') {
          normalizedList.push(normalizeSheetRecord(s as RawSheetRecord, `Sheet${idx + 1}`, idx));
        }
      });
    } else if (obj.sheets && typeof obj.sheets === 'object') {
      Object.entries(obj.sheets as Record<string, unknown>).forEach(([key, val], idx) => {
        if (val && typeof val === 'object') {
          normalizedList.push(normalizeSheetRecord(val as RawSheetRecord, key, idx));
        }
      });
    }
  } else if (Array.isArray(data)) {
    data.forEach((s, idx) => {
      if (s && typeof s === 'object') {
        normalizedList.push(normalizeSheetRecord(s as RawSheetRecord, `Sheet${idx + 1}`, idx));
      }
    });
  }

  normalizedList.sort((a, b) => a.tabOrder - b.tabOrder);

  const sheets: Record<string, WorkspaceSheetState> = {};
  const sheetNames: string[] = [];
  for (const sheet of normalizedList) {
    sheets[sheet.name] = sheet;
    sheetNames.push(sheet.name);
  }

  return {
    workspace,
    sheets,
    sheetNames,
    lastSavedAt
  };
}

export async function createWorkspaceSheetRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetName: string,
  headers?: string[]
): Promise<WorkspaceSheetState> {
  const validated = validateWorksheetName(sheetName);
  if (!validated.valid) {
    throw new StoreRpcError(400, 'INVALID_SHEET_NAME', validated.error);
  }

  const rpcArgs: Record<string, unknown> = {
    p_workspace: workspace,
    p_sheet_name: validated.name
  };
  if (headers && Array.isArray(headers) && headers.length > 0) {
    rpcArgs.p_headers = headers;
  }

  const { data, error } = await supabase.rpc('create_workspace_sheet', rpcArgs);

  if (error) {
    throw mapRpcError(error);
  }

  const raw = (data && typeof data === 'object' ? data : {}) as RawSheetRecord;
  return normalizeSheetRecord(
    {
      ...raw,
      name: raw.name ?? validated.name,
      headers: raw.headers ?? headers ?? []
    },
    validated.name,
    Number(raw.tab_order ?? raw.tabOrder ?? 0)
  );
}

export async function patchWorkspaceCellRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetId: string,
  edit: CellEditInput
): Promise<PatchCellResult> {
  const { data, error } = await supabase.rpc('patch_workspace_cell', {
    p_workspace: workspace,
    p_sheet_id: sheetId,
    p_row_index: edit.rowIndex,
    p_col_index: edit.colIndex,
    p_new_value: edit.value,
    p_expected_cell_version: edit.expectedCellVersion
  });

  if (error) {
    throw mapRpcError(error);
  }

  const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
  const newCellVersion = Number(res.newCellVersion ?? res.new_cell_version ?? edit.expectedCellVersion + 1);
  const newRowVersionRaw = res.newRowVersion ?? res.new_row_version;

  return {
    sheetId: String(res.sheetId ?? res.sheet_id ?? sheetId),
    rowIndex: Number(res.rowIndex ?? res.row_index ?? edit.rowIndex),
    colIndex: Number(res.colIndex ?? res.col_index ?? edit.colIndex),
    newCellVersion,
    ...(newRowVersionRaw !== undefined ? { newRowVersion: Number(newRowVersionRaw) } : {}),
    ...(typeof (res.updatedAt ?? res.updated_at) === 'string'
      ? { updatedAt: String(res.updatedAt ?? res.updated_at) }
      : {})
  };
}

export async function patchWorkspaceCellsBatchRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetId: string,
  edits: CellEditInput[]
): Promise<PatchCellsBatchResult> {
  const { data, error } = await supabase.rpc('patch_workspace_cells_batch', {
    p_workspace: workspace,
    p_sheet_id: sheetId,
    p_edits: edits
  });

  if (error) {
    throw mapRpcError(error);
  }

  const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
  const rawUpdated = Array.isArray(res.updatedCells ?? res.updated_cells)
    ? ((res.updatedCells ?? res.updated_cells) as Record<string, unknown>[])
    : [];

  const updatedCells = rawUpdated.map((item, idx) => {
    const fallback = edits[idx];
    const rowVerRaw = item.newRowVersion ?? item.new_row_version;
    return {
      rowIndex: Number(item.rowIndex ?? item.row_index ?? fallback?.rowIndex ?? 0),
      colIndex: Number(item.colIndex ?? item.col_index ?? fallback?.colIndex ?? 0),
      newCellVersion: Number(
        item.newCellVersion ?? item.new_cell_version ?? (fallback ? fallback.expectedCellVersion + 1 : 1)
      ),
      ...(rowVerRaw !== undefined ? { newRowVersion: Number(rowVerRaw) } : {})
    };
  });

  return {
    sheetId: String(res.sheetId ?? res.sheet_id ?? sheetId),
    editCount: Number(res.editCount ?? res.edit_count ?? edits.length),
    updatedCells,
    ...(typeof (res.updatedAt ?? res.updated_at) === 'string'
      ? { updatedAt: String(res.updatedAt ?? res.updated_at) }
      : {})
  };
}

export async function patchWorkspaceHeaderRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetId: string,
  colIndex: number,
  newHeader: string,
  expectedHeaderVersion: number
): Promise<PatchHeaderResult> {
  const { data, error } = await supabase.rpc('patch_workspace_header', {
    p_workspace: workspace,
    p_sheet_id: sheetId,
    p_col_index: colIndex,
    p_new_header: newHeader,
    p_expected_header_version: expectedHeaderVersion
  });

  if (error) {
    throw mapRpcError(error);
  }

  const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
  const headers = Array.isArray(res.headers)
    ? res.headers.map((h) => (typeof h === 'string' ? h : String(h ?? '')))
    : [];

  return {
    sheetId: String(res.sheetId ?? res.sheet_id ?? sheetId),
    colIndex: Number(res.colIndex ?? res.col_index ?? colIndex),
    headers,
    newHeaderVersion: Number(res.newHeaderVersion ?? res.new_header_version ?? expectedHeaderVersion + 1),
    ...(typeof (res.updatedAt ?? res.updated_at) === 'string'
      ? { updatedAt: String(res.updatedAt ?? res.updated_at) }
      : {})
  };
}

export interface InitialImportResult {
  status: 'imported' | 'already_imported';
  alreadyImported: boolean;
  importBatchId: string | null;
  receptionSheets: number;
  receptionRows: number;
  doctorSheets: number;
  doctorRows: number;
  totalSheets: number;
  totalRows: number;
}

export async function importInitialWorkbooksRpc(
  supabase: SupabaseClient,
  batchKey: string,
  payload: unknown
): Promise<InitialImportResult> {
  const { data, error } = await supabase.rpc('import_initial_workbooks', {
    p_batch_key: batchKey,
    p_payload: payload
  });

  if (error) {
    throw mapRpcError(error);
  }

  const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
  const alreadyImported = Boolean(res.already_imported ?? res.alreadyImported);
  const receptionSheets = Number(res.reception_sheets ?? res.receptionSheets ?? 0);
  const receptionRows = Number(res.reception_rows ?? res.receptionRows ?? 0);
  const doctorSheets = Number(res.doctor_sheets ?? res.doctorSheets ?? 0);
  const doctorRows = Number(res.doctor_rows ?? res.doctorRows ?? 0);

  return {
    status: alreadyImported ? 'already_imported' : 'imported',
    alreadyImported,
    importBatchId: res.import_batch_id || res.importBatchId ? String(res.import_batch_id ?? res.importBatchId) : null,
    receptionSheets,
    receptionRows,
    doctorSheets,
    doctorRows,
    totalSheets: Number(res.total_sheets ?? res.totalSheets ?? receptionSheets + doctorSheets),
    totalRows: Number(res.total_rows ?? res.totalRows ?? receptionRows + doctorRows)
  };
}

export async function renameWorkspaceSheetRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetId: string,
  newName: string
): Promise<{ sheetId: string; name: string }> {
  const validated = validateWorksheetName(newName);
  if (!validated.valid) {
    throw new StoreRpcError(400, 'INVALID_SHEET_NAME', validated.error);
  }

  const { data, error } = await supabase.rpc('rename_workspace_sheet', {
    p_workspace: workspace,
    p_sheet_id: sheetId,
    p_new_name: validated.name
  });

  if (!error) {
    const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
    return {
      sheetId: String(res.sheetId ?? sheetId),
      name: String(res.name ?? validated.name)
    };
  }

  // Graceful fallback to direct table update if RPC is not yet registered in Supabase
  if (
    error.code === 'PGRST202' ||
    /function public\.rename_workspace_sheet.*does not exist/i.test(error.message || '')
  ) {
    const { data: updateData, error: updateError } = await supabase
      .from('workbook_sheets')
      .update({ name: validated.name, sheet_key: validated.name })
      .eq('id', sheetId)
      .eq('workspace', workspace)
      .select('id, name')
      .single();

    if (updateError) {
      throw mapRpcError(updateError);
    }
    return {
      sheetId: String(updateData.id),
      name: String(updateData.name)
    };
  }

  throw mapRpcError(error);
}

export async function deleteWorkspaceSheetRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetId: string
): Promise<{ deletedSheetId: string }> {
  const { data, error } = await supabase.rpc('delete_workspace_sheet', {
    p_workspace: workspace,
    p_sheet_id: sheetId
  });

  if (!error) {
    const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
    return {
      deletedSheetId: String(res.deletedSheetId ?? sheetId)
    };
  }

  // Graceful fallback to direct table delete if RPC is not yet registered in Supabase
  if (
    error.code === 'PGRST202' ||
    /function public\.delete_workspace_sheet.*does not exist/i.test(error.message || '')
  ) {
    const { count, error: countError } = await supabase
      .from('workbook_sheets')
      .select('id', { count: 'exact', head: true })
      .eq('workspace', workspace);

    if (countError) throw mapRpcError(countError);
    if ((count ?? 0) <= 1) {
      throw new StoreRpcError(
        400,
        'BAD_REQUEST',
        `Cannot delete the only remaining worksheet in the ${workspace} workspace.`
      );
    }

    const { error: rowsDeleteError } = await supabase
      .from('workbook_rows')
      .delete()
      .eq('sheet_id', sheetId);
    if (rowsDeleteError) throw mapRpcError(rowsDeleteError);

    const { error: sheetDeleteError } = await supabase
      .from('workbook_sheets')
      .delete()
      .eq('id', sheetId)
      .eq('workspace', workspace);
    if (sheetDeleteError) throw mapRpcError(sheetDeleteError);

    return { deletedSheetId: sheetId };
  }

  throw mapRpcError(error);
}

export async function reorderWorkspaceSheetsRpc(
  supabase: SupabaseClient,
  workspace: WorkspaceType,
  sheetIds: string[]
): Promise<{ reorderedCount: number }> {
  const { data, error } = await supabase.rpc('reorder_workspace_sheets', {
    p_workspace: workspace,
    p_sheet_ids: sheetIds
  });

  if (!error) {
    const res = (data && typeof data === 'object' ? data : {}) as Record<string, unknown>;
    return {
      reorderedCount: Number(res.reorderedCount ?? sheetIds.length)
    };
  }

  // Graceful fallback to direct table updates if RPC is not yet registered in Supabase
  if (
    error.code === 'PGRST202' ||
    /function public\.reorder_workspace_sheets.*does not exist/i.test(error.message || '')
  ) {
    await Promise.all(
      sheetIds.map((id, idx) =>
        supabase
          .from('workbook_sheets')
          .update({ tab_order: idx })
          .eq('id', id)
          .eq('workspace', workspace)
      )
    );
    return { reorderedCount: sheetIds.length };
  }

  throw mapRpcError(error);
}

