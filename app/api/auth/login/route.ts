import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { verifyActiveStaff } from '@/lib/auth';
import type { UserRole } from '@/lib/types';

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const VALID_ROLES = new Set<UserRole>(['admin', 'doctor', 'reception']);

export async function POST(req: NextRequest) {
  try {
    let body: { email?: unknown; password?: unknown; role?: unknown };
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
    const requestedRole = typeof body.role === 'string' ? (body.role.trim().toLowerCase() as UserRole) : null;

    if (!email || !password || !requestedRole) {
      return NextResponse.json(
        { authenticated: false, error: 'Staff email, password, and role are required.' },
        { status: 400 }
      );
    }

    if (!EMAIL_REGEX.test(email)) {
      return NextResponse.json(
        { authenticated: false, error: 'Please enter a valid staff email address.' },
        { status: 400 }
      );
    }

    if (!VALID_ROLES.has(requestedRole)) {
      return NextResponse.json(
        { authenticated: false, error: 'Invalid role selection.' },
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
      await supabase.auth.signOut();
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
      await supabase.auth.signOut();
      return NextResponse.json(
        { authenticated: false, error: verification.error },
        { status: 500 }
      );
    }

    // STRICT ROLE MATCHING CHECK
    const actualRole = verification.user.role;
    if (actualRole !== requestedRole) {
      await supabase.auth.signOut();
      return NextResponse.json(
        {
          authenticated: false,
          error: `Role mismatch: You attempted to sign in as ${requestedRole}, but your assigned staff profile role is ${actualRole}. Access denied.`
        },
        { status: 403 }
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
