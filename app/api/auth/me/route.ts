import { NextRequest, NextResponse } from 'next/server';
import { verifySessionToken } from '@/lib/auth';

export async function GET(req: NextRequest) {
  const sessionCookie = req.cookies.get('mk_session')?.value;

  if (!sessionCookie) {
    return NextResponse.json({ authenticated: false }, { status: 401 });
  }

  const payload = verifySessionToken(sessionCookie);
  if (!payload) {
    return NextResponse.json({ authenticated: false, error: 'Session expired or invalid' }, { status: 401 });
  }

  return NextResponse.json({
    authenticated: true,
    user: {
      id: payload.userId,
      username: payload.username,
      name: payload.name,
      role: payload.role,
      title: payload.title
    }
  });
}
