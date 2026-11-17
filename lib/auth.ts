import type { SupabaseClient } from '@supabase/supabase-js';
import { createClient } from '@/lib/supabase/server';
import type { User, UserRole } from '@/lib/types';

export interface StaffProfileRow {
  user_id: string;
  role: UserRole;
  display_name: string;
  is_active: boolean;
}

export interface StaffUser extends User {
  id: string;
  email: string;
  role: UserRole;
  title: string;
  displayName: string;
}

export type StaffVerificationResult =
  | { status: 'authenticated'; user: StaffUser }
  | { status: 'unauthenticated'; user: null }
  | { status: 'unauthorized'; user: null; error: string }
  | { status: 'error'; user: null; error: string };

const VALID_ROLES = new Set<UserRole>(['admin', 'doctor', 'reception']);

function isValidStaffRole(role: unknown): role is UserRole {
  return typeof role === 'string' && VALID_ROLES.has(role as UserRole);
}

/**
 * Verifies the current Supabase Auth session user against public.staff_profiles.
 * Rejects and signs out any account without an active, role-valid staff profile.
 */
export async function verifyActiveStaff(
  supabaseClient?: SupabaseClient
): Promise<StaffVerificationResult> {
  const supabase = supabaseClient ?? (await createClient());

  const {
    data: { user: authUser },
    error: authError
  } = await supabase.auth.getUser();

  if (authError || !authUser) {
    return { status: 'unauthenticated', user: null };
  }

  const { data: profile, error: profileError } = await supabase
    .from('staff_profiles')
    .select('user_id, role, display_name, is_active')
    .eq('user_id', authUser.id)
    .maybeSingle<StaffProfileRow>();

  if (profileError) {
    await supabase.auth.signOut();
    return {
      status: 'error',
      user: null,
      error: 'Failed to verify staff authorization profile.'
    };
  }

  if (!profile || profile.is_active !== true || !isValidStaffRole(profile.role)) {
    await supabase.auth.signOut();
    return {
      status: 'unauthorized',
      user: null,
      error: !profile
        ? 'Account is not registered as an authorized MK Eye Center staff member.'
        : 'Staff account is inactive or lacks an authorized role.'
    };
  }

  const email = authUser.email ?? '';
  const displayName = (profile.display_name ?? '').trim() || email;

  return {
    status: 'authenticated',
    user: {
      id: authUser.id,
      email,
      username: email || authUser.id,
      name: displayName,
      role: profile.role,
      title: displayName,
      displayName
    }
  };
}

