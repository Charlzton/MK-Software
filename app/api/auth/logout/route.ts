import { NextResponse } from 'next/server';
import { createClient } from '@/lib/supabase/server';

export async function POST() {
  try {
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
