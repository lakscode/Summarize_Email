import { createClient, type SupabaseClient, type User } from "jsr:@supabase/supabase-js@2";

export function env(name: string, fallback?: string): string {
  const value = Deno.env.get(name) ?? fallback;
  if (value === undefined) throw new Error(`Missing server setting ${name}`);
  return value;
}

// Client that runs queries as the caller, so RLS applies.
export function userClient(authHeader: string): SupabaseClient {
  return createClient(env("SUPABASE_URL"), env("SUPABASE_ANON_KEY"), {
    global: { headers: { Authorization: authHeader } },
  });
}

// Service-role client that bypasses RLS. Server-side only.
export function adminClient(): SupabaseClient {
  return createClient(env("SUPABASE_URL"), env("SUPABASE_SERVICE_ROLE_KEY"), {
    auth: { persistSession: false },
  });
}

export async function callerUser(req: Request): Promise<{ user: User; sb: SupabaseClient } | null> {
  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return null;
  const sb = userClient(authHeader);
  const { data: { user }, error } = await sb.auth.getUser(authHeader.replace(/^Bearer /i, ""));
  return error || !user ? null : { user, sb };
}
