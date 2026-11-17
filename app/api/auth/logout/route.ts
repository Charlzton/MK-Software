import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';
import { isSupabaseConfigured } from '@/lib/supabase/config';
import { clearMockSessionCookie } from '@/lib/mockStore';

export async function POST(req: NextRequest) {
  try {
    if (!isSupabaseConfigured()) {
      const res = NextResponse.json({
        success: true,
        message: 'Logged out successfully'
      });
      clearMockSessionCookie(res, req);
      return res;
    }

    const supabase = await createClient();
    const { error } = await supabase.auth.signOut();

    if (error) {
      return NextResponse.json(
        { success: false, error: error.message || 'Failed to sign out.' },
        { status: 500 }
      );
    }

    return NextResponse.json({
      success: true,
      message: 'Logged out successfully'
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Logout failed';
    return NextResponse.json({ success: false, error: msg }, { status: 500 });
  }
}
