import { NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { verifyActiveStaff } from '@/lib/auth';

const NO_STORE_HEADERS = {
  'Cache-Control': 'private, no-cache, no-store, must-revalidate, max-age=0'
};

export async function GET() {
  try {
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

    return NextResponse.json(
      {
        authenticated: true,
        user: verification.user
      },
      { headers: NO_STORE_HEADERS }
    );
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Session verification error';
    return NextResponse.json({ authenticated: false, error: msg }, { status: 500, headers: NO_STORE_HEADERS });
  }
}
