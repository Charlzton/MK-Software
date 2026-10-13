import crypto from 'crypto';

export interface SyncResult {
  success: boolean;
  message: string;
  timestamp: string;
  details?: Record<string, unknown>;
  sheets?: Record<string, { name: string; headers: string[]; rows: (string | number)[][] }>;
}

export interface CellUpdatePayload {
  row: number; // 1-indexed
  col: number; // 1-indexed
  value: string | number;
}

/**
 * Computes the HMAC-SHA256 cryptographic signature matching Google Apps Script.
 */
function createSignedPayload(action: string, extraData: Record<string, unknown> = {}): string {
  const timestamp = Date.now();
  const secret = process.env.GOOGLE_SHEETS_SHARED_SECRET || '';

  const bodyData: Record<string, unknown> = {
    action,
    timestamp,
    ...extraData
  };

  // If secret is configured, compute HMAC signature
  if (secret) {
    const messageToSign = `${timestamp}:${JSON.stringify(bodyData)}`;
    const signature = crypto.createHmac('sha256', secret).update(messageToSign).digest('hex');
    bodyData.signature = signature;
  }

  return JSON.stringify(bodyData);
}

/**
 * Executes a secured POST request to the Google Apps Script Web App.
 */
export async function callAppsScript(
  action: string,
  extraData: Record<string, unknown> = {},
  timeoutMs = 15000,
  overrideUrl?: string
): Promise<SyncResult> {
  const scriptUrl = overrideUrl || process.env.GOOGLE_APPS_SCRIPT_URL;

  if (!scriptUrl || scriptUrl.trim() === '') {
    return {
      success: true,
      message: 'Local server persistent mode active (Google Apps Script URL not configured).',
      timestamp: new Date().toISOString()
    };
  }

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

    const postBody = createSignedPayload(action, extraData);

    const res = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: postBody,
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!res.ok) {
      return {
        success: false,
        message: `Google Sheets endpoint returned HTTP ${res.status} (${res.statusText})`,
        timestamp: new Date().toISOString()
      };
    }

    const data = await res.json().catch(() => ({ status: 'unknown' }));

    if (data.status === 'error') {
      return {
        success: false,
        message: data.message || 'Error reported by Google Sheets backend.',
        timestamp: new Date().toISOString(),
        details: data
      };
    }

    return {
      success: true,
      message: data.message || 'Operation successfully executed in Google Sheets.',
      timestamp: new Date().toISOString(),
      details: data,
      sheets: data.sheets
    };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      success: false,
      message: `Failed to communicate with Google Sheets: ${msg}`,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Reads the entire live workbook from Google Sheets.
 */
export async function fetchWorkbookFromGoogle(
  workspace: string,
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('GET_WORKBOOK', { workspace, spreadsheetId }, 20000, overrideUrl);
}

/**
 * Reads a single specific worksheet from Google Sheets.
 */
export async function fetchSheetFromGoogle(
  workspace: string,
  sheetName: string,
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('GET_SHEET', { workspace, sheet: sheetName, spreadsheetId }, 15000, overrideUrl);
}

/**
 * Commits the entire workbook to Google Sheets.
 */
export async function syncEntireWorkbookToGoogle(
  workspace: string,
  sheets: Record<string, any>,
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('SAVE_WORKBOOK', { workspace, sheets, spreadsheetId }, 25000, overrideUrl);
}

/**
 * Commits rows for a single worksheet to Google Sheets.
 */
export async function syncSheetToGoogle(
  workspace: string,
  sheetName: string,
  headers: string[],
  rows: (string | number)[][],
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('UPDATE_SHEET', { workspace, sheet: sheetName, headers, rows, spreadsheetId }, 20000, overrideUrl);
}

/**
 * Creates a brand new worksheet tab in Google Sheets.
 */
export async function createWorksheetInGoogle(
  workspace: string,
  sheetName: string,
  headers: string[] = [],
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('CREATE_WORKSHEET', { workspace, sheetName, headers, spreadsheetId }, 20000, overrideUrl);
}

/**
 * Performs targeted updates on specific cell coordinates.
 */
export async function updateSpecificCellsInGoogle(
  workspace: string,
  sheetName: string,
  updates: CellUpdatePayload[],
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('UPDATE_CELLS', { workspace, sheet: sheetName, updates, spreadsheetId }, 15000, overrideUrl);
}

/**
 * Pings Google Apps Script for connectivity & spreadsheet metadata.
 */
export async function pingGoogleSheetsBackend(
  workspace: string,
  spreadsheetId?: string,
  overrideUrl?: string
): Promise<SyncResult> {
  return callAppsScript('PING', { workspace, spreadsheetId }, 10000, overrideUrl);
}
