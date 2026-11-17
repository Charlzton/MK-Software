import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { verifyActiveStaff } from '@/lib/auth';
import {
  StoreRpcError,
  importInitialWorkbooksRpc,
  loadWorkspaceWorkbook
} from '@/lib/store';
import { prepareInitialWorkbookImport } from '@/lib/importWorkbook';

const NO_STORE_HEADERS = {
  'Cache-Control': 'private, no-cache, no-store, must-revalidate, max-age=0'
};

function handleImportError(err: unknown) {
  if (err instanceof StoreRpcError) {
    return NextResponse.json(
      {
        error: err.message,
        code: err.code,
        ...(err.details ? { details: err.details } : {})
      },
      { status: err.status, headers: NO_STORE_HEADERS }
    );
  }
  const msg = err instanceof Error ? err.message : 'Unexpected import error.';
  return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE_HEADERS });
}

function countWorkbookRows(sheets: Record<string, { rows: unknown[][] }>): number {
  let count = 0;
  for (const sheet of Object.values(sheets)) {
    count += Array.isArray(sheet.rows) ? sheet.rows.length : 0;
  }
  return count;
}

export async function GET() {
  try {
    const supabase = await createClient();
    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthenticated') {
      return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401, headers: NO_STORE_HEADERS });
    }
    if (verification.status === 'unauthorized') {
      return NextResponse.json({ error: verification.error }, { status: 403, headers: NO_STORE_HEADERS });
    }
    if (verification.status === 'error') {
      return NextResponse.json({ error: verification.error }, { status: 500, headers: NO_STORE_HEADERS });
    }

    if (verification.user.role !== 'admin') {
      return NextResponse.json(
        { error: 'Forbidden: Only administrator accounts can access initial workbook import.' },
        { status: 403, headers: NO_STORE_HEADERS }
      );
    }

    const [receptionWb, doctorWb] = await Promise.all([
      loadWorkspaceWorkbook(supabase, 'reception'),
      loadWorkspaceWorkbook(supabase, 'doctor')
    ]);

    const receptionSheetCount = receptionWb.sheetNames.length;
    const receptionRowCount = countWorkbookRows(receptionWb.sheets);
    const doctorSheetCount = doctorWb.sheetNames.length;
    const doctorRowCount = countWorkbookRows(doctorWb.sheets);

    const prepared = prepareInitialWorkbookImport();

    return NextResponse.json(
      {
        canImport: receptionSheetCount === 0 && doctorSheetCount === 0,
        currentCounts: {
          receptionSheets: receptionSheetCount,
          receptionRows: receptionRowCount,
          doctorSheets: doctorSheetCount,
          doctorRows: doctorRowCount
        },
        expectedCounts: prepared.expectedCounts,
        batchKey: prepared.batchKey
      },
      { headers: NO_STORE_HEADERS }
    );
  } catch (err: unknown) {
    return handleImportError(err);
  }
}

export async function POST(req: NextRequest) {
  try {
    const supabase = await createClient();
    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthenticated') {
      return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401, headers: NO_STORE_HEADERS });
    }
    if (verification.status === 'unauthorized') {
      return NextResponse.json({ error: verification.error }, { status: 403, headers: NO_STORE_HEADERS });
    }
    if (verification.status === 'error') {
      return NextResponse.json({ error: verification.error }, { status: 500, headers: NO_STORE_HEADERS });
    }

    if (verification.user.role !== 'admin') {
      return NextResponse.json(
        { error: 'Forbidden: Only administrator accounts can initiate the initial workbook import.' },
        { status: 403, headers: NO_STORE_HEADERS }
      );
    }

    let body: { confirm?: unknown } = {};
    try {
      body = (await req.json()) as { confirm?: unknown };
    } catch {
      return NextResponse.json(
        { error: 'Invalid JSON request body.' },
        { status: 400, headers: NO_STORE_HEADERS }
      );
    }

    if (body.confirm !== true) {
      return NextResponse.json(
        { error: 'Explicit import confirmation (confirm: true) is required.' },
        { status: 400, headers: NO_STORE_HEADERS }
      );
    }

    const prepared = prepareInitialWorkbookImport();
    const result = await importInitialWorkbooksRpc(
      supabase,
      prepared.batchKey,
      prepared.payload
    );

    return NextResponse.json(
      {
        success: true,
        ...result
      },
      { headers: NO_STORE_HEADERS }
    );
  } catch (err: unknown) {
    return handleImportError(err);
  }
}
