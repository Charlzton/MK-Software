import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { verifyActiveStaff } from '@/lib/auth';

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export async function POST(req: NextRequest) {
  try {
    let body: { email?: unknown; password?: unknown };
    try {
      body = await req.json();
    } catch {
      return NextResponse.json(
        { authenticated: false, error: 'Invalid request body.' },
        { status: 400 }
      );
    }

    const email = typeof body.email === 'string' ? body.email.trim().toLowerCase() : '';
    const password = typeof body.password === 'string' ? body.password : '';

    if (!email || !password) {
      return NextResponse.json(
        { authenticated: false, error: 'Staff email and password are required.' },
        { status: 400 }
      );
    }

    if (!EMAIL_REGEX.test(email)) {
      return NextResponse.json(
        { authenticated: false, error: 'Please enter a valid staff email address.' },
        { status: 400 }
      );
    }

    const supabase = await createClient();

    const { error: signInError } = await supabase.auth.signInWithPassword({
      email,
      password
    });

    if (signInError) {
      return NextResponse.json(
        { authenticated: false, error: 'Invalid email or password.' },
        { status: 401 }
      );
    }

    const verification = await verifyActiveStaff(supabase);

    if (verification.status === 'unauthorized') {
      return NextResponse.json(
        { authenticated: false, error: verification.error },
        { status: 403 }
      );
    }

    if (verification.status === 'unauthenticated') {
      await supabase.auth.signOut();
      return NextResponse.json(
        { authenticated: false, error: 'Unable to establish authenticated session.' },
        { status: 401 }
      );
    }

    if (verification.status === 'error') {
      return NextResponse.json(
        { authenticated: false, error: verification.error },
        { status: 500 }
      );
    }

    return NextResponse.json(
      {
        authenticated: true,
        user: verification.user
      },
      {
        headers: {
          'Cache-Control': 'private, no-cache, no-store, must-revalidate, max-age=0'
        }
      }
    );
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Authentication failure';
    return NextResponse.json({ authenticated: false, error: msg }, { status: 500 });
  }
}
