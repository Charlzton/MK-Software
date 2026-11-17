import { createServerClient } from '@supabase/ssr';
import { NextResponse, type NextRequest } from 'next/server';

function isHttpsRequest(request: NextRequest): boolean {
  const forwardedProto = request.headers.get('x-forwarded-proto')?.split(',')[0]?.trim().toLowerCase() || '';
  const forwarded = (request.headers.get('forwarded') || '').toLowerCase();
  const origin = (request.headers.get('origin') || '').toLowerCase();
  const referer = (request.headers.get('referer') || '').toLowerCase();
  const host = (request.headers.get('x-forwarded-host') || request.headers.get('host') || '').toLowerCase();

  return (
    request.nextUrl.protocol === 'https:' ||
    forwardedProto === 'https' ||
    forwarded.includes('proto=https') ||
    origin.startsWith('https://') ||
    referer.startsWith('https://') ||
    host.endsWith('.run.app')
  );
}

export async function updateSession(request: NextRequest) {
  let supabaseResponse = NextResponse.next({
    request
  });

  const pathname = request.nextUrl.pathname;
  if (pathname.startsWith('/api/')) {
    return supabaseResponse;
  }

  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey =
    process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ||
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

  // Do not crash existing app routes if Supabase environment variables are not yet configured
  if (!supabaseUrl || !supabaseKey) {
    return supabaseResponse;
  }

  const useIframeCookieOptions = isHttpsRequest(request);
  const iframeCookieOptions = useIframeCookieOptions
    ? {
        sameSite: 'none' as const,
        secure: true,
        partitioned: true
      }
    : undefined;

  const supabase = createServerClient(supabaseUrl, supabaseKey, {
    ...(iframeCookieOptions ? { cookieOptions: iframeCookieOptions } : {}),
    cookies: {
      getAll() {
        return request.cookies.getAll();
      },
      setAll(cookiesToSet) {
        cookiesToSet.forEach(({ name, value }) => {
          request.cookies.set(name, value);
        });
        supabaseResponse = NextResponse.next({
          request
        });
        cookiesToSet.forEach(({ name, value, options }) => {
          const sessionCookieOptions = { ...options };
          delete sessionCookieOptions.maxAge;
          delete sessionCookieOptions.expires;
          supabaseResponse.cookies.set(name, value, {
            ...sessionCookieOptions,
            ...(iframeCookieOptions || {})
          });
        });
      }
    }
  });

  // Refresh the Supabase Auth session when needed
  await supabase.auth.getUser();

  return supabaseResponse;
}
