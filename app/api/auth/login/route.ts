import { NextRequest, NextResponse } from 'next/server';
import { authenticateStaff, createSessionToken } from '@/lib/auth';

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { username, password } = body;

    if (!username || !password) {
      return NextResponse.json(
        { error: 'Staff username and password are required.' },
        { status: 400 }
      );
    }

    const user = authenticateStaff(username, password);

    if (!user) {
      // Intentional uniform rejection for security
      return NextResponse.json(
        { error: 'Invalid username or password.' },
        { status: 401 }
      );
    }

    const token = createSessionToken(user);

    const response = NextResponse.json({
      success: true,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role,
        title: user.title
      }
    });

    // Set secure HTTP-only cookie
    response.cookies.set('mk_session', token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'lax',
      path: '/',
      maxAge: 60 * 60 * 12 // 12 hours
    });

    return response;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Authentication failure';
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
