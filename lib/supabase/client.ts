import { createBrowserClient } from '@supabase/ssr';
import { isSupabaseConfigured } from '@/lib/supabase/config';

export function createClient() {
  if (!isSupabaseConfigured()) {
    return null;
  }

  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!;
  const supabaseKey =
    (process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ||
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY)!;

  const useIframeCookieOptions =
    typeof window !== 'undefined' &&
    (window.location.protocol === 'https:' || window.location.hostname.endsWith('.run.app'));

  return createBrowserClient(supabaseUrl, supabaseKey, {
    ...(useIframeCookieOptions
      ? {
          cookieOptions: {
            sameSite: 'none',
            secure: true,
            partitioned: true
          }
        }
      : {})
  });
}
