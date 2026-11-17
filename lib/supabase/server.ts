import { createServerClient } from '@supabase/ssr';
import { cookies, headers } from 'next/headers';

function isHttpsProxyContext(headerStore: Headers): boolean {
  const forwardedProto = headerStore.get('x-forwarded-proto')?.split(',')[0]?.trim().toLowerCase() || '';
  const forwarded = (headerStore.get('forwarded') || '').toLowerCase();
  const origin = (headerStore.get('origin') || '').toLowerCase();
  const referer = (headerStore.get('referer') || '').toLowerCase();
  const host = (headerStore.get('x-forwarded-host') || headerStore.get('host') || '').toLowerCase();

  return (
    forwardedProto === 'https' ||
    forwarded.includes('proto=https') ||
    origin.startsWith('https://') ||
    referer.startsWith('https://') ||
    host.endsWith('.run.app')
  );
}

export async function createClient() {
  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey =
    process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ||
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  if (!supabaseUrl || !supabaseKey) {
    throw new Error(
      'Missing Supabase environment variables: NEXT_PUBLIC_SUPABASE_URL and NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY (or NEXT_PUBLIC_SUPABASE_ANON_KEY).'
    );
  }

  const cookieStore = await cookies();
  const headerStore = await headers();
  const useIframeCookieOptions = isHttpsProxyContext(headerStore);
  const iframeCookieOptions = useIframeCookieOptions
    ? {
        sameSite: 'none' as const,
        secure: true,
        partitioned: true
      }
    : undefined;

  return createServerClient(supabaseUrl, supabaseKey, {
    ...(iframeCookieOptions ? { cookieOptions: iframeCookieOptions } : {}),
    cookies: {
      getAll() {
        return cookieStore.getAll();
      },
      setAll(cookiesToSet) {
        try {
          cookiesToSet.forEach(({ name, value, options }) => {
            const sessionCookieOptions = { ...options };
            delete sessionCookieOptions.maxAge;
            delete sessionCookieOptions.expires;
            cookieStore.set(name, value, {
              ...sessionCookieOptions,
              ...(iframeCookieOptions || {})
            });
          });
        } catch {
          // Called from a Server Component where cookie mutation is not permitted.
          // Session cookie updates are handled by Next.js middleware instead.
        }
      }
    }
  });
}
