import { NextRequest, NextResponse } from 'next/server';
import { verifySessionToken } from '@/lib/auth';
import {
  WorkspaceType,
  getWorkspaceSheets,
  getWorkspaceSheet,
  updateWorkspaceWorkbook,
  createWorkspaceSheet,
  setFullWorkbookFromRemote,
  isDoctorConfigured,
  getDoctorSpreadsheetId,
  getLastSavedTime
} from '@/lib/store';
import {
  syncEntireWorkbookToGoogle,
  syncSheetToGoogle,
  fetchWorkbookFromGoogle,
  createWorksheetInGoogle
} from '@/lib/sheets';
import { SheetDefinition } from '@/lib/data/rawWorkbookData';

function getAuthenticatedUser(req: NextRequest) {
  const cookie = req.cookies.get('mk_session')?.value;
  if (!cookie) return null;
  return verifySessionToken(cookie);
}

export async function GET(req: NextRequest) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401 });
  }

  const { searchParams } = new URL(req.url);
  const workspaceParam = searchParams.get('workspace');
  if (workspaceParam !== 'reception' && workspaceParam !== 'doctor') {
    return NextResponse.json({ error: 'Invalid or missing workspace parameter.' }, { status: 400 });
  }
  const workspace: WorkspaceType = workspaceParam;
  const sheetParam = searchParams.get('sheet');

  if (workspace === 'reception') {
    const receptionSpreadsheetId = process.env.GOOGLE_SPREADSHEET_ID || '11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI';

    if (process.env.GOOGLE_APPS_SCRIPT_URL) {
      try {
        const googleResult = await fetchWorkbookFromGoogle('reception', receptionSpreadsheetId);
        if (googleResult.success && googleResult.sheets) {
          setFullWorkbookFromRemote('reception', googleResult.sheets as Record<string, SheetDefinition>);
        }
      } catch (err) {
        console.warn('Could not fetch latest Reception Google Sheets data, using local store:', err);
      }
    }

    if (sheetParam) {
      const sheetData = getWorkspaceSheet('reception', sheetParam);
      if (!sheetData) {
        return NextResponse.json({ error: 'Worksheet not found' }, { status: 404 });
      }
      return NextResponse.json({
        workspace: 'reception',
        sheet: sheetData,
        lastSavedAt: getLastSavedTime('reception')
      });
    }

    const allSheets = getWorkspaceSheets('reception');
    return NextResponse.json({
      workspace: 'reception',
      sheets: allSheets,
      sheetNames: Object.keys(allSheets),
      isConfigured: true,
      spreadsheetId: receptionSpreadsheetId,
      lastSavedAt: getLastSavedTime('reception')
    });
  }

  // Doctor workspace
  const doctorConfigured = isDoctorConfigured();
  const doctorSpreadsheetId = getDoctorSpreadsheetId();

  if (doctorConfigured && process.env.GOOGLE_APPS_SCRIPT_URL) {
    try {
      const googleResult = await fetchWorkbookFromGoogle('doctor', doctorSpreadsheetId);
      if (googleResult.success && googleResult.sheets) {
        setFullWorkbookFromRemote('doctor', googleResult.sheets as Record<string, SheetDefinition>);
      }
    } catch (err) {
      console.warn('Could not fetch Doctor Google Sheets data:', err);
    }
  }

  if (sheetParam) {
    const sheetData = getWorkspaceSheet('doctor', sheetParam);
    if (!sheetData) {
      return NextResponse.json({ error: 'Worksheet not found' }, { status: 404 });
    }
    return NextResponse.json({
      workspace: 'doctor',
      sheet: sheetData,
      isConfigured: doctorConfigured,
      lastSavedAt: getLastSavedTime('doctor')
    });
  }

  const doctorSheets = getWorkspaceSheets('doctor');
  return NextResponse.json({
    workspace: 'doctor',
    sheets: doctorSheets,
    sheetNames: Object.keys(doctorSheets),
    isConfigured: doctorConfigured,
    spreadsheetId: doctorSpreadsheetId,
    configMessage: doctorConfigured
      ? 'Doctor Google Sheets workbook is connected.'
      : "Doctor's Google Sheets workbook is not yet configured. The doctor's original Excel workbook and its structure will be connected using DOCTOR_SPREADSHEET_ID.",
    lastSavedAt: getLastSavedTime('doctor')
  });
}

export async function POST(req: NextRequest) {
  const user = getAuthenticatedUser(req);
  if (!user) {
    return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401 });
  }

  try {
    const body = await req.json();
    const workspaceParam = body.workspace;
    if (workspaceParam !== 'reception' && workspaceParam !== 'doctor') {
      return NextResponse.json({ error: 'Invalid or missing workspace parameter in request body.' }, { status: 400 });
    }
    const workspace: WorkspaceType = workspaceParam;
    const action = body.action || 'save';

    // 1. Action: CREATE_WORKSHEET
    if (action === 'create_sheet') {
      const sheetName = String(body.sheetName || '').trim();
      if (!sheetName) {
        return NextResponse.json({ error: 'Worksheet name cannot be empty.' }, { status: 400 });
      }

      const result = createWorkspaceSheet(workspace, sheetName, body.headers);
      if (!result.success) {
        return NextResponse.json({ error: result.error || 'Failed to create worksheet.' }, { status: 400 });
      }

      let googleSyncSuccess = true;
      let googleSyncMessage = 'Created in local workspace.';

      if (workspace === 'reception') {
        const receptionSpreadsheetId = process.env.GOOGLE_SPREADSHEET_ID || '11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI';
        if (process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await createWorksheetInGoogle(workspace, sheetName, body.headers || [], receptionSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        }
      } else if (workspace === 'doctor') {
        const doctorSpreadsheetId = getDoctorSpreadsheetId();
        if (isDoctorConfigured() && process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await createWorksheetInGoogle(workspace, sheetName, body.headers || [], doctorSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        }
      }

      const updatedSheets = getWorkspaceSheets(workspace);
      return NextResponse.json({
        success: true,
        workspace,
        createdSheet: sheetName,
        sheets: updatedSheets,
        sheetNames: Object.keys(updatedSheets),
        googleSyncSuccess,
        message: googleSyncSuccess ? 'Worksheet created successfully.' : `Created (${googleSyncMessage})`
      });
    }

    // 2. Action: SAVE_WORKBOOK
    if (body.sheets && typeof body.sheets === 'object') {
      const sheetsToSave = body.sheets as Record<string, (string | number)[][]>;
      updateWorkspaceWorkbook(workspace, sheetsToSave);

      let googleSyncMessage = 'Saved locally.';
      let googleSyncSuccess = true;

      if (workspace === 'reception') {
        const receptionSpreadsheetId = process.env.GOOGLE_SPREADSHEET_ID || '11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI';
        if (process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await syncEntireWorkbookToGoogle(workspace, sheetsToSave, receptionSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        }
      } else if (workspace === 'doctor') {
        const doctorSpreadsheetId = getDoctorSpreadsheetId();
        if (isDoctorConfigured() && process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await syncEntireWorkbookToGoogle(workspace, sheetsToSave, doctorSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        } else {
          googleSyncMessage = 'Saved to Doctor workspace. Google Sheets workbook not connected yet.';
        }
      }

      return NextResponse.json({
        success: true,
        workspace,
        message: googleSyncSuccess ? 'Saved' : `Saved locally (${googleSyncMessage})`,
        googleSyncSuccess,
        lastSavedAt: getLastSavedTime(workspace)
      });
    }

    // 3. Action: SAVE SINGLE SHEET
    if (body.sheet && body.rows && Array.isArray(body.rows)) {
      const { sheet, rows, headers } = body;
      updateWorkspaceWorkbook(workspace, { [sheet]: { headers: headers || [], rows } });

      let googleSyncMessage = 'Saved locally.';
      let googleSyncSuccess = true;

      if (workspace === 'reception') {
        const receptionSpreadsheetId = process.env.GOOGLE_SPREADSHEET_ID || '11rBhezcklOndIPa7YAHr3zIBh6WCXkp2235slmirnlI';
        if (process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await syncSheetToGoogle(workspace, sheet, headers || [], rows, receptionSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        }
      } else if (workspace === 'doctor') {
        const doctorSpreadsheetId = getDoctorSpreadsheetId();
        if (isDoctorConfigured() && process.env.GOOGLE_APPS_SCRIPT_URL) {
          const syncRes = await syncSheetToGoogle(workspace, sheet, headers || [], rows, doctorSpreadsheetId);
          googleSyncSuccess = syncRes.success;
          googleSyncMessage = syncRes.message;
        } else {
          googleSyncMessage = 'Saved to Doctor workspace.';
        }
      }

      return NextResponse.json({
        success: true,
        workspace,
        message: googleSyncSuccess ? 'Saved' : `Saved locally (${googleSyncMessage})`,
        googleSyncSuccess,
        lastSavedAt: getLastSavedTime(workspace)
      });
    }

    return NextResponse.json({ error: 'Invalid save payload.' }, { status: 400 });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Save error';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
