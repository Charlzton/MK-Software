import * as XLSX from 'xlsx';
import type { WorkspaceType, WorkspaceSheetState } from '@/lib/store';

export interface ParsedImportWorkbook {
  fileName: string;
  fileSize: number;
  sheetNames: string[];
  sheets: Record<string, {
    name: string;
    headers: string[];
    rows: (string | number)[][];
    columnCount: number;
    rowCount: number;
    isEmpty: boolean;
    warnings: string[];
  }>;
}

export type ImportMode = 'create' | 'append' | 'replace';

export interface SheetImportConfig {
  sheetName: string;
  mode: ImportMode;
  targetExistingSheetName?: string;
  columnMappingConfirmed?: boolean;
}

/**
 * Validates and parses an uploaded Excel (.xlsx, .xls) or CSV file using SheetJS.
 */
export async function parseUploadedWorkbook(file: File): Promise<ParsedImportWorkbook> {
  const ext = file.name.split('.').pop()?.toLowerCase() || '';
  if (!['xlsx', 'xls', 'csv', 'ods', 'txt'].includes(ext)) {
    throw new Error('Unsupported file format. Please upload an Excel (.xlsx, .xls) or CSV file.');
  }

  if (file.size > 25 * 1024 * 1024) {
    throw new Error('File size exceeds the 25MB limit.');
  }

  const buffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellNF: true, cellText: false });
  } catch (err) {
    throw new Error(
      `Failed to parse workbook file: ${err instanceof Error ? err.message : 'Corrupted or invalid file format.'}`
    );
  }

  const sheetNames = workbook.SheetNames;
  if (!sheetNames || sheetNames.length === 0) {
    throw new Error('The uploaded workbook contains no worksheets.');
  }

  const sheets: ParsedImportWorkbook['sheets'] = {};

  for (const sheetName of sheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' });
    const warnings: string[] = [];

    if (aoa.length === 0) {
      sheets[sheetName] = {
        name: sheetName,
        headers: [],
        rows: [],
        columnCount: 0,
        rowCount: 0,
        isEmpty: true,
        warnings: ['Worksheet is empty.']
      };
      continue;
    }

    // Extract headers from row 0
    const rawHeaders = aoa[0] || [];
    const headers = rawHeaders.map((h, i) => {
      const str = h !== null && h !== undefined ? String(h).trim() : '';
      return str !== '' ? str : `Column${i + 1}`;
    });

    const rawRows = aoa.slice(1);
    const rows: (string | number)[][] = [];

    for (let rIdx = 0; rIdx < rawRows.length; rIdx++) {
      const row = rawRows[rIdx];
      if (!Array.isArray(row)) continue;

      const normalizedRow = row.map((cell) => {
        if (cell === null || cell === undefined) return '';
        if (typeof cell === 'number' && Number.isFinite(cell)) return cell;
        if (cell instanceof Date) {
          return cell.toISOString().slice(0, 10);
        }
        const str = String(cell).trim();
        if (str.startsWith('=')) {
          warnings.push(`Row ${rIdx + 2} contains formula cells. Cached formula values will be used.`);
        }
        const num = Number(str);
        if (str !== '' && !Number.isNaN(num) && /^-?\d+(\.\d+)?$/.test(str)) {
          return num;
        }
        return str;
      });

      // Skip entirely empty trailing rows
      if (normalizedRow.some((c) => c !== '')) {
        rows.push(normalizedRow);
      }
    }

    const columnCount = Math.max(headers.length, ...rows.map((r) => r.length), 0);
    const rowCount = rows.length;

    sheets[sheetName] = {
      name: sheetName,
      headers,
      rows,
      columnCount,
      rowCount,
      isEmpty: rowCount === 0 && headers.length === 0,
      warnings: Array.from(new Set(warnings))
    };
  }

  return {
    fileName: file.name,
    fileSize: file.size,
    sheetNames: Object.keys(sheets),
    sheets
  };
}

/**
 * Generates and downloads a professional Excel export file (.xlsx).
 */
export function exportSheetsToExcel(
  workspace: WorkspaceType,
  sheetsToExport: Array<{ name: string; headers: string[]; rows: (string | number)[][] }>,
  exportTypeLabel: string
) {
  const wb = XLSX.utils.book_new();

  for (const s of sheetsToExport) {
    const aoa = [s.headers, ...s.rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Apply basic column width estimations
    const colWidths = s.headers.map((h, cIdx) => {
      let maxLen = String(h || '').length;
      for (const row of s.rows) {
        const val = row[cIdx];
        if (val !== undefined && val !== null) {
          const len = String(val).length;
          if (len > maxLen) maxLen = len;
        }
      }
      return { wch: Math.min(Math.max(maxLen + 3, 12), 45) };
    });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, s.name);
  }

  const today = new Date().toISOString().slice(0, 10);
  const workspaceTitle = workspace === 'doctor' ? 'Doctor' : 'Reception';
  const filename = `MK_Eye_Center_${workspaceTitle}_${exportTypeLabel}_${today}.xlsx`;

  XLSX.writeFile(wb, filename);
}
