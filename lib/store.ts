import fs from 'fs';
import path from 'path';
import { SheetDefinition, INITIAL_WORKBOOK, SHEET_NAMES } from './data/rawWorkbookData';
import { INITIAL_DOCTOR_WORKBOOK, DOCTOR_SHEET_NAMES } from './data/rawDoctorWorkbookData';

export type WorkspaceType = 'reception' | 'doctor';

interface WorkbookDatabase {
  sheets: Record<string, SheetDefinition>;
  lastSavedAt: string;
}

const DATA_DIR = path.join(process.cwd(), 'data');
const RECEPTION_DB_FILE = path.join(DATA_DIR, 'reception_workbook_store.json');
const LEGACY_DB_FILE = path.join(DATA_DIR, 'mk_workbook_store.json');
const DOCTOR_DB_FILE = path.join(DATA_DIR, 'doctor_workbook_store.json');

const DEFAULT_BLANK_HEADERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'];

function ensureDirectoryExists() {
  if (!fs.existsSync(DATA_DIR)) {
    try {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    } catch {
      // ignore
    }
  }
}

function ensureReceptionStore(): WorkbookDatabase {
  ensureDirectoryExists();

  // Migrate legacy store file if present
  if (!fs.existsSync(RECEPTION_DB_FILE) && fs.existsSync(LEGACY_DB_FILE)) {
    try {
      const legacyContent = fs.readFileSync(LEGACY_DB_FILE, 'utf-8');
      fs.writeFileSync(RECEPTION_DB_FILE, legacyContent, 'utf-8');
    } catch {
      // ignore
    }
  }

  if (!fs.existsSync(RECEPTION_DB_FILE)) {
    const initialDb: WorkbookDatabase = {
      sheets: { ...INITIAL_WORKBOOK },
      lastSavedAt: new Date().toISOString()
    };
    try {
      fs.writeFileSync(RECEPTION_DB_FILE, JSON.stringify(initialDb, null, 2), 'utf-8');
      return initialDb;
    } catch (err) {
      console.error('Failed to create reception database file:', err);
      return initialDb;
    }
  }

  try {
    const content = fs.readFileSync(RECEPTION_DB_FILE, 'utf-8');
    const parsed = JSON.parse(content) as WorkbookDatabase;
    return parsed;
  } catch (err) {
    console.error('Failed to read reception database file, restoring defaults:', err);
    return {
      sheets: { ...INITIAL_WORKBOOK },
      lastSavedAt: new Date().toISOString()
    };
  }
}

function ensureDoctorStore(): WorkbookDatabase {
  ensureDirectoryExists();

  if (!fs.existsSync(DOCTOR_DB_FILE)) {
    const initialDb: WorkbookDatabase = {
      sheets: { ...INITIAL_DOCTOR_WORKBOOK },
      lastSavedAt: new Date().toISOString()
    };
    try {
      fs.writeFileSync(DOCTOR_DB_FILE, JSON.stringify(initialDb, null, 2), 'utf-8');
      return initialDb;
    } catch (err) {
      console.error('Failed to create doctor database file:', err);
      return initialDb;
    }
  }

  try {
    const content = fs.readFileSync(DOCTOR_DB_FILE, 'utf-8');
    const parsed = JSON.parse(content) as WorkbookDatabase;
    // If the existing file only has a single generic Sheet1, upgrade to real doctor worksheets
    if (parsed && parsed.sheets && (!parsed.sheets['Lens Rx'] || Object.keys(parsed.sheets).length <= 1)) {
      parsed.sheets = { ...INITIAL_DOCTOR_WORKBOOK, ...parsed.sheets };
      persistStore(DOCTOR_DB_FILE, parsed);
    }
    return parsed;
  } catch (err) {
    console.error('Failed to read doctor database file, restoring defaults:', err);
    return {
      sheets: { ...INITIAL_DOCTOR_WORKBOOK },
      lastSavedAt: new Date().toISOString()
    };
  }
}

function persistStore(file: string, db: WorkbookDatabase) {
  try {
    ensureDirectoryExists();
    fs.writeFileSync(file, JSON.stringify(db, null, 2), 'utf-8');
  } catch (err) {
    console.error(`Failed to persist store ${file}:`, err);
  }
}

const DEFAULT_DOCTOR_SPREADSHEET_ID = '1kdrkBEP2Iok8iKBRvBQdHalIcK5bl3uMrKgB46iWh94';

/**
 * Checks if the Doctor Google Sheets workbook has been configured.
 */
export function isDoctorConfigured(): boolean {
  return true;
}

export function getDoctorSpreadsheetId(): string {
  const envId = process.env.DOCTOR_SPREADSHEET_ID;
  if (envId && envId.trim() !== '' && envId !== 'undefined' && envId !== 'null') {
    return envId.trim();
  }
  return DEFAULT_DOCTOR_SPREADSHEET_ID;
}

export function getWorkspaceSheets(workspace: WorkspaceType): Record<string, SheetDefinition> {
  if (workspace === 'doctor') {
    const db = ensureDoctorStore();
    return db.sheets;
  }
  const db = ensureReceptionStore();
  return db.sheets;
}

export function getWorkspaceSheet(workspace: WorkspaceType, sheetName: string): SheetDefinition | null {
  const sheets = getWorkspaceSheets(workspace);
  return sheets[sheetName] || null;
}

export type SheetSaveInput =
  | (string | number)[][]
  | {
      headers?: string[];
      rows: (string | number)[][];
    };

export function updateWorkspaceWorkbook(
  workspace: WorkspaceType,
  sheets: Record<string, SheetSaveInput>
): boolean {
  if (workspace === 'doctor') {
    const db = ensureDoctorStore();
    for (const [name, input] of Object.entries(sheets)) {
      if (!input) continue;

      let headers: string[] | undefined;
      let rows: (string | number)[][];

      if (Array.isArray(input)) {
        rows = input;
      } else {
        headers = input.headers;
        rows = input.rows || [];
      }

      if (db.sheets[name]) {
        if (headers && Array.isArray(headers)) {
          db.sheets[name].headers = headers;
        }
        db.sheets[name].rows = rows;
      } else {
        // Newly created sheet
        db.sheets[name] = {
          name,
          headers: headers && Array.isArray(headers) ? headers : DEFAULT_BLANK_HEADERS,
          rows
        };
      }
    }
    db.lastSavedAt = new Date().toISOString();
    persistStore(DOCTOR_DB_FILE, db);
    return true;
  }

  // Reception
  const db = ensureReceptionStore();
  for (const [name, input] of Object.entries(sheets)) {
    if (!input) continue;

    let headers: string[] | undefined;
    let rows: (string | number)[][];

    if (Array.isArray(input)) {
      rows = input;
    } else {
      headers = input.headers;
      rows = input.rows || [];
    }

    if (db.sheets[name]) {
      if (headers && Array.isArray(headers)) {
        db.sheets[name].headers = headers;
      }
      db.sheets[name].rows = rows;
    } else if (INITIAL_WORKBOOK[name]) {
      db.sheets[name] = {
        ...INITIAL_WORKBOOK[name],
        ...(headers ? { headers } : {}),
        rows
      };
    } else {
      // User-created sheet in reception
      db.sheets[name] = {
        name,
        headers: headers && Array.isArray(headers) ? headers : DEFAULT_BLANK_HEADERS,
        rows
      };
    }
  }
  db.lastSavedAt = new Date().toISOString();
  persistStore(RECEPTION_DB_FILE, db);
  return true;
}

export function createWorkspaceSheet(
  workspace: WorkspaceType,
  sheetName: string,
  headers?: string[]
): { success: boolean; error?: string } {
  const trimmed = sheetName.trim();
  if (!trimmed) {
    return { success: false, error: 'Worksheet name cannot be empty.' };
  }

  // Check Excel illegal characters
  const illegalChars = /[\\/?*[\]:]/;
  if (illegalChars.test(trimmed)) {
    return {
      success: false,
      error: 'Worksheet name cannot contain any of the following characters: \\ / ? * [ ] :'
    };
  }

  if (trimmed.length > 31) {
    return {
      success: false,
      error: 'Worksheet name cannot exceed 31 characters.'
    };
  }

  if (workspace === 'doctor') {
    const db = ensureDoctorStore();
    const existing = Object.keys(db.sheets).find((k) => k.toLowerCase() === trimmed.toLowerCase());
    if (existing) {
      return { success: false, error: `A worksheet named "${existing}" already exists in the Doctor workspace.` };
    }

    db.sheets[trimmed] = {
      name: trimmed,
      headers: headers || DEFAULT_BLANK_HEADERS,
      rows: []
    };
    db.lastSavedAt = new Date().toISOString();
    persistStore(DOCTOR_DB_FILE, db);
    return { success: true };
  }

  // Reception
  const db = ensureReceptionStore();
  const existing = Object.keys(db.sheets).find((k) => k.toLowerCase() === trimmed.toLowerCase());
  if (existing) {
    return { success: false, error: `A worksheet named "${existing}" already exists in the Reception workspace.` };
  }

  const defaultReceptionHeaders = [
    'Date',
    'Recorded_By',
    'Phone_No',
    'Examined_By',
    'Name',
    'Date_of_Birth',
    'Test_Type',
    'Test_Fee',
    'Drops_No',
    'Medication',
    'Total',
    'Glass Total',
    'Paid',
    'Glasses BAL',
    'Review Date',
    'Review Test'
  ];

  db.sheets[trimmed] = {
    name: trimmed,
    headers: headers || defaultReceptionHeaders,
    rows: []
  };
  db.lastSavedAt = new Date().toISOString();
  persistStore(RECEPTION_DB_FILE, db);
  return { success: true };
}

export function setFullWorkbookFromRemote(
  workspace: WorkspaceType,
  sheets: Record<string, SheetDefinition>
): void {
  if (workspace === 'doctor') {
    const db = ensureDoctorStore();
    db.sheets = { ...db.sheets, ...sheets };
    db.lastSavedAt = new Date().toISOString();
    persistStore(DOCTOR_DB_FILE, db);
    return;
  }

  const db = ensureReceptionStore();
  db.sheets = { ...db.sheets, ...sheets };
  db.lastSavedAt = new Date().toISOString();
  persistStore(RECEPTION_DB_FILE, db);
}

export function getLastSavedTime(workspace: WorkspaceType = 'reception'): string {
  if (workspace === 'doctor') {
    const db = ensureDoctorStore();
    return db.lastSavedAt;
  }
  const db = ensureReceptionStore();
  return db.lastSavedAt;
}

// Backward-compatible exports for existing references
export function getAllSheets(): Record<string, SheetDefinition> {
  return getWorkspaceSheets('reception');
}

export function getSheet(sheetName: string): SheetDefinition | null {
  return getWorkspaceSheet('reception', sheetName);
}

export function updateWorkbook(sheets: Record<string, (string | number)[][]>): boolean {
  return updateWorkspaceWorkbook('reception', sheets);
}

export function updateSheetRows(sheetName: string, rows: (string | number)[][]): boolean {
  return updateWorkspaceWorkbook('reception', { [sheetName]: rows });
}
