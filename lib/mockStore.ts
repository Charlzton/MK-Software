import crypto from 'crypto';
import type { NextRequest, NextResponse } from 'next/server';
import type { StaffUser } from '@/lib/auth';
import type { UserRole } from '@/lib/types';
import {
  type CellEditInput,
  type InitialImportResult,
  type PatchCellResult,
  type PatchCellsBatchResult,
  type PatchHeaderResult,
  type WorkspaceSheetState,
  type WorkspaceType,
  type WorkspaceWorkbookResult,
  StoreRpcError,
  validateWorksheetName
} from '@/lib/store';
import { prepareInitialWorkbookImport } from '@/lib/importWorkbook';

export const MOCK_SESSION_COOKIE = 'mk_mock_session';

interface MockWorkbookState {
  workbooks: Record<WorkspaceType, Record<string, WorkspaceSheetState>>;
  lastSavedAt: Record<WorkspaceType, string | null>;
  importedBatchKey: string | null;
  nextSheetNum: number;
}

const globalForMock = globalThis as unknown as {
  __mkMockStore?: MockWorkbookState;
};

function createSeededState(): MockWorkbookState {
  const prepared = prepareInitialWorkbookImport();
  const now = new Date().toISOString();

  const buildSheets = (
    workspace: WorkspaceType,
    rawSheets: typeof prepared.payload.reception
  ): Record<string, WorkspaceSheetState> => {
    const sheets: Record<string, WorkspaceSheetState> = {};
    for (const raw of rawSheets) {
      const rowVersions: Record<number, number> = {};
      const cellVersions: Record<number, Record<number, number>> = {};
      const rowUpdatedAt: Record<number, string> = {};

      raw.rows.forEach((row, rIdx) => {
        rowVersions[rIdx] = 1;
        rowUpdatedAt[rIdx] = now;
        cellVersions[rIdx] = {};
        row.forEach((_, cIdx) => {
          cellVersions[rIdx][cIdx] = 1;
        });
      });

      sheets[raw.name] = {
        id: `mock-${workspace}-${raw.tabOrder}-${raw.sheetKey}`,
        name: raw.name,
        sheetKey: raw.sheetKey,
        tabOrder: raw.tabOrder,
        headers: [...raw.headers],
        headerVersion: 1,
        rows: raw.rows.map((r) => [...r]),
        rowVersions,
        cellVersions,
        rowUpdatedAt
      };
    }
    return sheets;
  };

  return {
    workbooks: {
      reception: buildSheets('reception', prepared.payload.reception),
      doctor: buildSheets('doctor', prepared.payload.doctor)
    },
    lastSavedAt: {
      reception: now,
      doctor: now
    },
    importedBatchKey: prepared.batchKey,
    nextSheetNum: 100
  };
}

function getMockStore(): MockWorkbookState {
  if (!globalForMock.__mkMockStore) {
    globalForMock.__mkMockStore = createSeededState();
  }
  return globalForMock.__mkMockStore;
}

function getSessionSecret(): string {
  return process.env.SESSION_SECRET || 'mk-eye-center-local-fallback-secret';
}

function signPayload(encoded: string): string {
  return crypto.createHmac('sha256', getSessionSecret()).update(encoded).digest('base64url');
}

export function createMockSessionToken(user: StaffUser): string {
  const encoded = Buffer.from(JSON.stringify(user), 'utf-8').toString('base64url');
  const sig = signPayload(encoded);
  return `${encoded}.${sig}`;
}

export function verifyMockSessionToken(token: string | undefined): StaffUser | null {
  if (!token || !token.includes('.')) return null;
  const [encoded, sig] = token.split('.');
  if (!encoded || !sig) return null;

  const expected = signPayload(encoded);
  if (sig !== expected) return null;

  try {
    const parsed = JSON.parse(Buffer.from(encoded, 'base64url').toString('utf-8')) as StaffUser;
    if (!parsed || !parsed.id || !parsed.role) return null;
    return parsed;
  } catch {
    return null;
  }
}

export function getMockUserFromRequest(req: NextRequest): StaffUser | null {
  const cookieVal = req.cookies.get(MOCK_SESSION_COOKIE)?.value;
  return verifyMockSessionToken(cookieVal);
}

export function isHttpsRequest(req: NextRequest): boolean {
  const forwardedProto = req.headers.get('x-forwarded-proto')?.split(',')[0]?.trim().toLowerCase() || '';
  const forwarded = (req.headers.get('forwarded') || '').toLowerCase();
  const origin = (req.headers.get('origin') || '').toLowerCase();
  const referer = (req.headers.get('referer') || '').toLowerCase();
  const host = (req.headers.get('x-forwarded-host') || req.headers.get('host') || '').toLowerCase();

  return (
    req.nextUrl.protocol === 'https:' ||
    forwardedProto === 'https' ||
    forwarded.includes('proto=https') ||
    origin.startsWith('https://') ||
    referer.startsWith('https://') ||
    host.endsWith('.run.app')
  );
}

export function setMockSessionCookie(res: NextResponse, req: NextRequest, user: StaffUser): void {
  const token = createMockSessionToken(user);
  const https = isHttpsRequest(req);
  res.cookies.set(MOCK_SESSION_COOKIE, token, {
    httpOnly: true,
    path: '/',
    maxAge: 60 * 60 * 24 * 7,
    ...(https ? { sameSite: 'none' as const, secure: true, partitioned: true } : { sameSite: 'lax' as const })
  });
}

export function clearMockSessionCookie(res: NextResponse, req: NextRequest): void {
  const https = isHttpsRequest(req);
  res.cookies.set(MOCK_SESSION_COOKIE, '', {
    httpOnly: true,
    path: '/',
    maxAge: 0,
    ...(https ? { sameSite: 'none' as const, secure: true, partitioned: true } : { sameSite: 'lax' as const })
  });
}

export function buildMockStaffUser(email: string, role: UserRole): StaffUser {
  const localPart = email.split('@')[0] || role;
  const displayName =
    role === 'admin'
      ? `Admin (${localPart})`
      : role === 'doctor'
        ? `Dr. ${localPart}`
        : `Reception (${localPart})`;

  return {
    id: `mock-${role}-${crypto.createHash('sha256').update(email).digest('hex').slice(0, 12)}`,
    email,
    username: email,
    name: displayName,
    role,
    title: displayName,
    displayName
  };
}

function findSheetById(
  sheets: Record<string, WorkspaceSheetState>,
  sheetId: string
): WorkspaceSheetState {
  for (const sheet of Object.values(sheets)) {
    if (sheet.id === sheetId) return sheet;
  }
  throw new StoreRpcError(404, 'NOT_FOUND', `Worksheet ${sheetId} not found.`);
}

export function mockLoadWorkspaceWorkbook(workspace: WorkspaceType): WorkspaceWorkbookResult {
  const store = getMockStore();
  const sheetsMap = store.workbooks[workspace];
  const sorted = Object.values(sheetsMap).sort((a, b) => a.tabOrder - b.tabOrder);

  const sheets: Record<string, WorkspaceSheetState> = {};
  const sheetNames: string[] = [];
  for (const sheet of sorted) {
    sheets[sheet.name] = sheet;
    sheetNames.push(sheet.name);
  }

  return {
    workspace,
    sheets,
    sheetNames,
    lastSavedAt: store.lastSavedAt[workspace]
  };
}

export function mockCreateWorkspaceSheet(
  workspace: WorkspaceType,
  sheetName: string,
  headers?: string[]
): WorkspaceSheetState {
  const validated = validateWorksheetName(sheetName);
  if (!validated.valid) {
    throw new StoreRpcError(400, 'INVALID_SHEET_NAME', validated.error);
  }

  const store = getMockStore();
  const sheetsMap = store.workbooks[workspace];

  if (Object.keys(sheetsMap).some((k) => k.toLowerCase() === validated.name.toLowerCase())) {
    throw new StoreRpcError(400, 'BAD_REQUEST', `Worksheet "${validated.name}" already exists.`);
  }

  const maxOrder = Object.values(sheetsMap).reduce((m, s) => Math.max(m, s.tabOrder), -1);
  const newId = `mock-${workspace}-${++store.nextSheetNum}`;
  const created: WorkspaceSheetState = {
    id: newId,
    name: validated.name,
    sheetKey: validated.name,
    tabOrder: maxOrder + 1,
    headers: headers && headers.length > 0 ? [...headers] : ['Date', 'Patient Name', 'Notes'],
    headerVersion: 1,
    rows: [],
    rowVersions: {},
    cellVersions: {},
    rowUpdatedAt: {}
  };

  sheetsMap[validated.name] = created;
  store.lastSavedAt[workspace] = new Date().toISOString();
  return created;
}

export function mockPatchWorkspaceCell(
  workspace: WorkspaceType,
  sheetId: string,
  edit: CellEditInput
): PatchCellResult {
  const store = getMockStore();
  const sheet = findSheetById(store.workbooks[workspace], sheetId);
  const currentCellVersion = sheet.cellVersions[edit.rowIndex]?.[edit.colIndex] ?? 0;

  if (edit.expectedCellVersion !== currentCellVersion) {
    throw new StoreRpcError(
      409,
      'VERSION_CONFLICT',
      'Version conflict: another user modified this data since you loaded it. Reload the latest data and retry your edit.'
    );
  }

  while (sheet.rows.length <= edit.rowIndex) {
    sheet.rows.push([]);
  }
  while (sheet.rows[edit.rowIndex].length <= edit.colIndex) {
    sheet.rows[edit.rowIndex].push('');
  }

  sheet.rows[edit.rowIndex][edit.colIndex] = edit.value;
  const newCellVersion = currentCellVersion + 1;
  if (!sheet.cellVersions[edit.rowIndex]) {
    sheet.cellVersions[edit.rowIndex] = {};
  }
  sheet.cellVersions[edit.rowIndex][edit.colIndex] = newCellVersion;

  const newRowVersion = (sheet.rowVersions[edit.rowIndex] ?? 0) + 1;
  sheet.rowVersions[edit.rowIndex] = newRowVersion;

  const now = new Date().toISOString();
  if (!sheet.rowUpdatedAt) sheet.rowUpdatedAt = {};
  sheet.rowUpdatedAt[edit.rowIndex] = now;
  store.lastSavedAt[workspace] = now;

  return {
    sheetId,
    rowIndex: edit.rowIndex,
    colIndex: edit.colIndex,
    newCellVersion,
    newRowVersion,
    updatedAt: now
  };
}

export function mockPatchWorkspaceCellsBatch(
  workspace: WorkspaceType,
  sheetId: string,
  edits: CellEditInput[]
): PatchCellsBatchResult {
  const store = getMockStore();
  const sheet = findSheetById(store.workbooks[workspace], sheetId);
  const now = new Date().toISOString();

  // Validate versions first
  for (const edit of edits) {
    const currentCellVersion = sheet.cellVersions[edit.rowIndex]?.[edit.colIndex] ?? 0;
    if (edit.expectedCellVersion !== currentCellVersion) {
      throw new StoreRpcError(
        409,
        'VERSION_CONFLICT',
        'Version conflict: another user modified this data since you loaded it. Reload the latest data and retry your edit.'
      );
    }
  }

  const updatedCells: PatchCellsBatchResult['updatedCells'] = [];
  for (const edit of edits) {
    while (sheet.rows.length <= edit.rowIndex) {
      sheet.rows.push([]);
    }
    while (sheet.rows[edit.rowIndex].length <= edit.colIndex) {
      sheet.rows[edit.rowIndex].push('');
    }
    sheet.rows[edit.rowIndex][edit.colIndex] = edit.value;

    const currentCellVersion = sheet.cellVersions[edit.rowIndex]?.[edit.colIndex] ?? 0;
    const newCellVersion = currentCellVersion + 1;
    if (!sheet.cellVersions[edit.rowIndex]) {
      sheet.cellVersions[edit.rowIndex] = {};
    }
    sheet.cellVersions[edit.rowIndex][edit.colIndex] = newCellVersion;

    const newRowVersion = (sheet.rowVersions[edit.rowIndex] ?? 0) + 1;
    sheet.rowVersions[edit.rowIndex] = newRowVersion;
    if (!sheet.rowUpdatedAt) sheet.rowUpdatedAt = {};
    sheet.rowUpdatedAt[edit.rowIndex] = now;

    updatedCells.push({
      rowIndex: edit.rowIndex,
      colIndex: edit.colIndex,
      newCellVersion,
      newRowVersion
    });
  }

  store.lastSavedAt[workspace] = now;
  return {
    sheetId,
    editCount: edits.length,
    updatedCells,
    updatedAt: now
  };
}

export function mockPatchWorkspaceHeader(
  workspace: WorkspaceType,
  sheetId: string,
  colIndex: number,
  newHeader: string,
  expectedHeaderVersion: number
): PatchHeaderResult {
  const store = getMockStore();
  const sheet = findSheetById(store.workbooks[workspace], sheetId);

  if (expectedHeaderVersion !== sheet.headerVersion) {
    throw new StoreRpcError(
      409,
      'VERSION_CONFLICT',
      'Version conflict: another user modified this data since you loaded it. Reload the latest data and retry your edit.'
    );
  }

  while (sheet.headers.length <= colIndex) {
    sheet.headers.push('');
  }
  sheet.headers[colIndex] = newHeader;
  sheet.headerVersion += 1;
  const now = new Date().toISOString();
  store.lastSavedAt[workspace] = now;

  return {
    sheetId,
    colIndex,
    headers: [...sheet.headers],
    newHeaderVersion: sheet.headerVersion,
    updatedAt: now
  };
}

export function mockRenameWorkspaceSheet(
  workspace: WorkspaceType,
  sheetId: string,
  newName: string
): { sheetId: string; name: string } {
  const validated = validateWorksheetName(newName);
  if (!validated.valid) {
    throw new StoreRpcError(400, 'INVALID_SHEET_NAME', validated.error);
  }

  const store = getMockStore();
  const sheetsMap = store.workbooks[workspace];
  const sheet = findSheetById(sheetsMap, sheetId);

  if (
    Object.values(sheetsMap).some(
      (s) => s.id !== sheetId && s.name.toLowerCase() === validated.name.toLowerCase()
    )
  ) {
    throw new StoreRpcError(400, 'BAD_REQUEST', `Worksheet "${validated.name}" already exists.`);
  }

  const oldName = sheet.name;
  if (oldName !== validated.name) {
    delete sheetsMap[oldName];
    sheet.name = validated.name;
    sheet.sheetKey = validated.name;
    sheetsMap[validated.name] = sheet;
  }

  store.lastSavedAt[workspace] = new Date().toISOString();
  return { sheetId, name: validated.name };
}

export function mockDeleteWorkspaceSheet(
  workspace: WorkspaceType,
  sheetId: string
): { deletedSheetId: string } {
  const store = getMockStore();
  const sheetsMap = store.workbooks[workspace];

  if (Object.keys(sheetsMap).length <= 1) {
    throw new StoreRpcError(
      400,
      'BAD_REQUEST',
      `Cannot delete the only remaining worksheet in the ${workspace} workspace.`
    );
  }

  const sheet = findSheetById(sheetsMap, sheetId);
  delete sheetsMap[sheet.name];

  Object.values(sheetsMap)
    .sort((a, b) => a.tabOrder - b.tabOrder)
    .forEach((s, idx) => {
      s.tabOrder = idx;
    });

  store.lastSavedAt[workspace] = new Date().toISOString();

  return { deletedSheetId: sheetId };
}

export function mockReorderWorkspaceSheets(
  workspace: WorkspaceType,
  sheetIds: string[]
): { reorderedCount: number } {
  const store = getMockStore();
  const sheetsMap = store.workbooks[workspace];
  const byId = new Map<string, WorkspaceSheetState>();
  for (const s of Object.values(sheetsMap)) {
    byId.set(s.id, s);
  }

  let count = 0;
  sheetIds.forEach((id, idx) => {
    const s = byId.get(id);
    if (s) {
      s.tabOrder = idx;
      count += 1;
    }
  });

  store.lastSavedAt[workspace] = new Date().toISOString();
  return { reorderedCount: count };
}

export function mockImportInitialWorkbooks(): InitialImportResult {
  const store = getMockStore();
  const prepared = prepareInitialWorkbookImport();

  return {
    status: 'already_imported',
    alreadyImported: true,
    importBatchId: store.importedBatchKey,
    receptionSheets: prepared.expectedCounts.receptionSheets,
    receptionRows: prepared.expectedCounts.receptionRows,
    doctorSheets: prepared.expectedCounts.doctorSheets,
    doctorRows: prepared.expectedCounts.doctorRows,
    totalSheets: prepared.expectedCounts.totalSheets,
    totalRows: prepared.expectedCounts.totalRows
  };
}
