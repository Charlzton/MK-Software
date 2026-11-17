import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { isSupabaseConfigured } from '@/lib/supabase/config';
import { getMockUserFromRequest, mockLoadWorkspaceWorkbook } from '@/lib/mockStore';
import { verifyActiveStaff } from '@/lib/auth';
import { loadWorkspaceWorkbook, type WorkspaceType } from '@/lib/store';

const NO_STORE_HEADERS = {
  'Cache-Control': 'private, no-cache, no-store, must-revalidate, max-age=0'
};

export async function GET(req: NextRequest) {
  try {
    const includeWorkbook = req.nextUrl.searchParams.get('includeWorkbook') === '1';

    if (!isSupabaseConfigured()) {
      const mockUser = getMockUserFromRequest(req);
      if (!mockUser) {
        return NextResponse.json({ authenticated: false }, { status: 401, headers: NO_STORE_HEADERS });
      }

      if (includeWorkbook) {
        const initialWorkspace: WorkspaceType = mockUser.role === 'doctor' ? 'doctor' : 'reception';
        const initialWorkbook = mockLoadWorkspaceWorkbook(initialWorkspace);
        return NextResponse.json(
          {
            authenticated: true,
            user: mockUser,
            initialWorkspace,
            initialWorkbook: {
              sheets: initialWorkbook.sheets,
              sheetNames: initialWorkbook.sheetNames,
              lastSavedAt: initialWorkbook.lastSavedAt
            }
          },
          { headers: NO_STORE_HEADERS }
        );
      }

      return NextResponse.json(
        { authenticated: true, user: mockUser },
        { headers: NO_STORE_HEADERS }
      );
    }

    const supabase = await createClient();
    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthenticated') {
      return NextResponse.json({ authenticated: false }, { status: 401, headers: NO_STORE_HEADERS });
    }

    if (verification.status === 'unauthorized') {
      return NextResponse.json(
        { authenticated: false, error: verification.error },
        { status: 403, headers: NO_STORE_HEADERS }
      );
    }

    if (verification.status === 'error') {
      return NextResponse.json(
        { authenticated: false, error: verification.error },
        { status: 500, headers: NO_STORE_HEADERS }
      );
    }

    let initialWorkspace: WorkspaceType | undefined;
    let initialWorkbookPayload: {
      sheets: Awaited<ReturnType<typeof loadWorkspaceWorkbook>>['sheets'];
      sheetNames: string[];
      lastSavedAt: string | null;
    } | null = null;

    if (includeWorkbook) {
      initialWorkspace = verification.user.role === 'doctor' ? 'doctor' : 'reception';
      try {
        const wb = await loadWorkspaceWorkbook(supabase, initialWorkspace);
        initialWorkbookPayload = {
          sheets: wb.sheets,
          sheetNames: wb.sheetNames,
          lastSavedAt: wb.lastSavedAt
        };
      } catch {
        // Non-fatal: client falls back to fetching /api/records?workspace=...
      }
    }

    return NextResponse.json(
      {
        authenticated: true,
        user: verification.user,
        ...(initialWorkspace ? { initialWorkspace } : {}),
        ...(initialWorkbookPayload ? { initialWorkbook: initialWorkbookPayload } : {})
      },
      { headers: NO_STORE_HEADERS }
    );
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Session verification error';
    return NextResponse.json({ authenticated: false, error: msg }, { status: 500, headers: NO_STORE_HEADERS });
  }
}
