import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { verifyActiveStaff } from '@/lib/auth';
import {
  WorkspaceType,
  CellEditInput,
  StoreRpcError,
  canRoleAccessWorkspace,
  loadWorkspaceWorkbook,
  createWorkspaceSheetRpc,
  renameWorkspaceSheetRpc,
  deleteWorkspaceSheetRpc,
  reorderWorkspaceSheetsRpc,
  patchWorkspaceCellRpc,
  patchWorkspaceCellsBatchRpc,
  patchWorkspaceHeaderRpc
} from '@/lib/store';

function isValidWorkspace(val: unknown): val is WorkspaceType {
  return val === 'reception' || val === 'doctor';
}

function isNonNegativeInt(val: unknown): val is number {
  return typeof val === 'number' && Number.isInteger(val) && val >= 0;
}

function isValidCellScalar(val: unknown): val is string | number {
  return typeof val === 'string' || (typeof val === 'number' && Number.isFinite(val));
}

function handleErrorResponse(err: unknown) {
  if (err instanceof StoreRpcError) {
    return NextResponse.json(
      {
        error: err.message,
        code: err.code,
        ...(err.details ? { details: err.details } : {})
      },
      { status: err.status }
    );
  }
  const msg = err instanceof Error ? err.message : 'Unexpected server error';
  return NextResponse.json({ error: msg }, { status: 500 });
}

export async function GET(req: NextRequest) {
  try {
    const supabase = await createClient();
    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthenticated') {
      return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401 });
    }
    if (verification.status === 'unauthorized') {
      return NextResponse.json({ error: verification.error }, { status: 403 });
    }
    if (verification.status === 'error') {
      return NextResponse.json({ error: verification.error }, { status: 500 });
    }

    const staff = verification.user;
    const { searchParams } = new URL(req.url);
    const workspaceParam = searchParams.get('workspace');

    if (!isValidWorkspace(workspaceParam)) {
      return NextResponse.json({ error: 'Invalid or missing workspace parameter.' }, { status: 400 });
    }

    if (!canRoleAccessWorkspace(staff.role, workspaceParam)) {
      return NextResponse.json(
        { error: `Forbidden: ${staff.role} role cannot access the ${workspaceParam} workspace.` },
        { status: 403 }
      );
    }

    const workbook = await loadWorkspaceWorkbook(supabase, workspaceParam);
    const sheetParam = searchParams.get('sheet');

    if (sheetParam) {
      const sheetData = workbook.sheets[sheetParam];
      if (!sheetData) {
        return NextResponse.json({ error: 'Worksheet not found' }, { status: 404 });
      }
      return NextResponse.json({
        workspace: workspaceParam,
        sheet: sheetData,
        lastSavedAt: workbook.lastSavedAt
      });
    }

    return NextResponse.json({
      workspace: workspaceParam,
      sheets: workbook.sheets,
      sheetNames: workbook.sheetNames,
      isConfigured: true,
      lastSavedAt: workbook.lastSavedAt
    });
  } catch (err: unknown) {
    return handleErrorResponse(err);
  }
}

export async function POST(req: NextRequest) {
  try {
    const supabase = await createClient();
    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthenticated') {
      return NextResponse.json({ error: 'Unauthorized. Please sign in.' }, { status: 401 });
    }
    if (verification.status === 'unauthorized') {
      return NextResponse.json({ error: verification.error }, { status: 403 });
    }
    if (verification.status === 'error') {
      return NextResponse.json({ error: verification.error }, { status: 500 });
    }

    const staff = verification.user;

    let body: Record<string, unknown>;
    try {
      body = (await req.json()) as Record<string, unknown>;
    } catch {
      return NextResponse.json({ error: 'Invalid JSON request body.' }, { status: 400 });
    }

    const workspaceParam = body.workspace;
    if (!isValidWorkspace(workspaceParam)) {
      return NextResponse.json(
        { error: 'Invalid or missing workspace parameter in request body.' },
        { status: 400 }
      );
    }

    if (!canRoleAccessWorkspace(staff.role, workspaceParam)) {
      return NextResponse.json(
        { error: `Forbidden: ${staff.role} role cannot modify the ${workspaceParam} workspace.` },
        { status: 403 }
      );
    }

    const action = typeof body.action === 'string' ? body.action : '';

    // 1. Create new worksheet
    if (action === 'create_sheet') {
      const sheetName = typeof body.sheetName === 'string' ? body.sheetName.trim() : '';
      if (!sheetName) {
        return NextResponse.json({ error: 'Worksheet name cannot be empty.' }, { status: 400 });
      }

      let headers: string[] | undefined;
      if (body.headers !== undefined) {
        if (!Array.isArray(body.headers) || !body.headers.every((h) => typeof h === 'string')) {
          return NextResponse.json({ error: 'headers must be an array of strings.' }, { status: 400 });
        }
        headers = body.headers;
      }

      const created = await createWorkspaceSheetRpc(supabase, workspaceParam, sheetName, headers);
      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        createdSheet: created.name,
        sheet: created,
        message: 'Worksheet created successfully.'
      });
    }

    // 2. Patch single cell with mandatory optimistic concurrency version
    if (action === 'patch_cell') {
      const sheetId = typeof body.sheetId === 'string' ? body.sheetId.trim() : '';
      const rowIndex = body.rowIndex;
      const colIndex = body.colIndex;
      const value = body.value;
      const expectedCellVersion = body.expectedCellVersion;

      if (!sheetId) {
        return NextResponse.json({ error: 'sheetId is required.' }, { status: 400 });
      }
      if (!isNonNegativeInt(rowIndex) || rowIndex >= 100000) {
        return NextResponse.json({ error: 'rowIndex must be an integer between 0 and 99999.' }, { status: 400 });
      }
      if (!isNonNegativeInt(colIndex) || colIndex >= 256) {
        return NextResponse.json({ error: 'colIndex must be an integer between 0 and 255.' }, { status: 400 });
      }
      if (!isValidCellScalar(value)) {
        return NextResponse.json({ error: 'value must be a string or finite number.' }, { status: 400 });
      }
      if (!isNonNegativeInt(expectedCellVersion)) {
        return NextResponse.json(
          { error: 'expectedCellVersion is required and must be a non-negative integer.' },
          { status: 400 }
        );
      }

      const result = await patchWorkspaceCellRpc(supabase, workspaceParam, sheetId, {
        rowIndex,
        colIndex,
        value,
        expectedCellVersion
      });

      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    // 3. Patch batch of cells atomically with mandatory expectedCellVersion per cell
    if (action === 'patch_cells_batch') {
      const sheetId = typeof body.sheetId === 'string' ? body.sheetId.trim() : '';
      if (!sheetId) {
        return NextResponse.json({ error: 'sheetId is required.' }, { status: 400 });
      }

      if (!Array.isArray(body.edits) || body.edits.length === 0 || body.edits.length > 500) {
        return NextResponse.json(
          { error: 'edits must be a non-empty array of 1..500 cell edits.' },
          { status: 400 }
        );
      }

      const validatedEdits: CellEditInput[] = [];
      for (let i = 0; i < body.edits.length; i++) {
        const item = body.edits[i] as Record<string, unknown> | null;
        if (!item || typeof item !== 'object') {
          return NextResponse.json({ error: `Invalid edit object at index ${i}.` }, { status: 400 });
        }
        if (!isNonNegativeInt(item.rowIndex) || item.rowIndex >= 100000) {
          return NextResponse.json({ error: `Invalid rowIndex at edit index ${i}.` }, { status: 400 });
        }
        if (!isNonNegativeInt(item.colIndex) || item.colIndex >= 256) {
          return NextResponse.json({ error: `Invalid colIndex at edit index ${i}.` }, { status: 400 });
        }
        if (!isValidCellScalar(item.value)) {
          return NextResponse.json({ error: `Invalid cell value at edit index ${i}.` }, { status: 400 });
        }
        if (!isNonNegativeInt(item.expectedCellVersion)) {
          return NextResponse.json(
            { error: `Missing or invalid expectedCellVersion at edit index ${i}.` },
            { status: 400 }
          );
        }
        validatedEdits.push({
          rowIndex: item.rowIndex,
          colIndex: item.colIndex,
          value: item.value,
          expectedCellVersion: item.expectedCellVersion
        });
      }

      const result = await patchWorkspaceCellsBatchRpc(supabase, workspaceParam, sheetId, validatedEdits);

      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    // 4. Patch header cell with mandatory expectedHeaderVersion
    if (action === 'patch_header') {
      const sheetId = typeof body.sheetId === 'string' ? body.sheetId.trim() : '';
      const colIndex = body.colIndex;
      const newHeader = body.newHeader;
      const expectedHeaderVersion = body.expectedHeaderVersion;

      if (!sheetId) {
        return NextResponse.json({ error: 'sheetId is required.' }, { status: 400 });
      }
      if (!isNonNegativeInt(colIndex) || colIndex >= 256) {
        return NextResponse.json({ error: 'colIndex must be an integer between 0 and 255.' }, { status: 400 });
      }
      if (typeof newHeader !== 'string') {
        return NextResponse.json({ error: 'newHeader must be a string.' }, { status: 400 });
      }
      if (!isNonNegativeInt(expectedHeaderVersion) || expectedHeaderVersion < 1) {
        return NextResponse.json(
          { error: 'expectedHeaderVersion is required and must be >= 1.' },
          { status: 400 }
        );
      }

      const result = await patchWorkspaceHeaderRpc(
        supabase,
        workspaceParam,
        sheetId,
        colIndex,
        newHeader,
        expectedHeaderVersion
      );

      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    // 5. Rename worksheet
    if (action === 'rename_sheet') {
      const sheetId = typeof body.sheetId === 'string' ? body.sheetId.trim() : '';
      const newName = typeof body.newName === 'string' ? body.newName.trim() : '';

      if (!sheetId) {
        return NextResponse.json({ error: 'sheetId is required.' }, { status: 400 });
      }
      if (!newName) {
        return NextResponse.json({ error: 'newName is required.' }, { status: 400 });
      }

      const result = await renameWorkspaceSheetRpc(supabase, workspaceParam, sheetId, newName);
      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    // 6. Delete worksheet
    if (action === 'delete_sheet') {
      const sheetId = typeof body.sheetId === 'string' ? body.sheetId.trim() : '';
      if (!sheetId) {
        return NextResponse.json({ error: 'sheetId is required.' }, { status: 400 });
      }

      const result = await deleteWorkspaceSheetRpc(supabase, workspaceParam, sheetId);
      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    // 7. Reorder worksheets
    if (action === 'reorder_sheets') {
      if (!Array.isArray(body.sheetIds) || body.sheetIds.length === 0 || !body.sheetIds.every((id) => typeof id === 'string')) {
        return NextResponse.json({ error: 'sheetIds must be a non-empty array of strings.' }, { status: 400 });
      }

      const result = await reorderWorkspaceSheetsRpc(supabase, workspaceParam, body.sheetIds as string[]);
      return NextResponse.json({
        success: true,
        workspace: workspaceParam,
        ...result
      });
    }

    return NextResponse.json(
      {
        error:
          'Invalid or unsupported action. Supported actions: create_sheet, rename_sheet, delete_sheet, reorder_sheets, patch_cell, patch_cells_batch, patch_header.'
      },
      { status: 400 }
    );
  } catch (err: unknown) {
    return handleErrorResponse(err);
  }
}
