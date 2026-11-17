import { createHash } from 'crypto';
import { INITIAL_WORKBOOK, SHEET_NAMES, type SheetDefinition } from '@/lib/data/rawWorkbookData';
import { INITIAL_DOCTOR_WORKBOOK, DOCTOR_SHEET_NAMES } from '@/lib/data/rawDoctorWorkbookData';
import type { WorkspaceType } from '@/lib/store';

export interface ImportSheetPayload {
  name: string;
  sheetKey: string;
  tabOrder: number;
  headers: string[];
  rows: (string | number)[][];
}

export interface InitialImportPayload {
  reception: ImportSheetPayload[];
  doctor: ImportSheetPayload[];
}

export interface ExpectedImportCounts {
  receptionSheets: number;
  receptionRows: number;
  doctorSheets: number;
  doctorRows: number;
  totalSheets: number;
  totalRows: number;
}

export interface PreparedInitialImport {
  batchKey: string;
  checksum: string;
  payload: InitialImportPayload;
  expectedCounts: ExpectedImportCounts;
}

function makeSheetKey(workspace: WorkspaceType, tabOrder: number, name: string): string {
  const slug = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'sheet';
  return `${workspace}__${tabOrder}__${slug}`;
}

function buildWorkspaceImportSheets(
  workspace: WorkspaceType,
  orderedNames: string[],
  workbook: Record<string, SheetDefinition>
): { sheets: ImportSheetPayload[]; totalRows: number } {
  const sheets: ImportSheetPayload[] = [];
  let totalRows = 0;

  orderedNames.forEach((sheetName, tabOrder) => {
    const def = workbook[sheetName];
    if (!def) {
      throw new Error(`Missing seed worksheet definition for ${workspace}:${sheetName}`);
    }

    // Preserve exact headers (including blank strings & duplicate column names) and exact rows (including ragged rows & blanks)
    const headers = def.headers.map((h) => String(h));
    const rows = def.rows.map((row) =>
      row.map((cell) => {
        if (typeof cell === 'number' && Number.isFinite(cell)) return cell;
        if (typeof cell === 'string') return cell;
        if (cell === null || cell === undefined) return '';
        return String(cell);
      })
    );

    totalRows += rows.length;
    sheets.push({
      name: def.name,
      sheetKey: makeSheetKey(workspace, tabOrder, def.name),
      tabOrder,
      headers,
      rows
    });
  });

  return { sheets, totalRows };
}

export function prepareInitialWorkbookImport(): PreparedInitialImport {
  const reception = buildWorkspaceImportSheets('reception', SHEET_NAMES, INITIAL_WORKBOOK);
  const doctor = buildWorkspaceImportSheets('doctor', DOCTOR_SHEET_NAMES, INITIAL_DOCTOR_WORKBOOK);

  const payload: InitialImportPayload = {
    reception: reception.sheets,
    doctor: doctor.sheets
  };

  const checksum = createHash('sha256').update(JSON.stringify(payload)).digest('hex').slice(0, 16);
  const batchKey = `mk_eye_center_initial_import_v1_${checksum}`;

  return {
    batchKey,
    checksum,
    payload,
    expectedCounts: {
      receptionSheets: reception.sheets.length,
      receptionRows: reception.totalRows,
      doctorSheets: doctor.sheets.length,
      doctorRows: doctor.totalRows,
      totalSheets: reception.sheets.length + doctor.sheets.length,
      totalRows: reception.totalRows + doctor.totalRows
    }
  };
}
