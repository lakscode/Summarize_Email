// GET  /functions/v1/profile -> the signed-in user's profile
// PATCH /functions/v1/profile { full_name } -> update it
// Queries run as the caller (their JWT is forwarded), so RLS on public.profiles applies.
import { createClient } from "jsr:@supabase/supabase-js@2";
import { corsHeaders, json } from "../_shared/cors.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return json({ error: "Missing Authorization header" }, 401);

  const sb = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, {
    global: { headers: { Authorization: authHeader } },
  });

  const { data: { user }, error: userError } = await sb.auth.getUser(authHeader.replace(/^Bearer /i, ""));
  if (userError || !user) return json({ error: "Not signed in" }, 401);

  if (req.method === "GET") {
    const { data, error } = await sb
      .from("profiles")
      .select("id, email, full_name, created_at, updated_at")
      .eq("id", user.id)
      .maybeSingle();
    if (error) return json({ error: error.message }, 500);
    return json({ profile: data });
  }

  if (req.method === "PATCH") {
    const body = await req.json().catch(() => ({}));
    const fullName = typeof body.full_name === "string" ? body.full_name.trim() : "";
    if (!fullName) return json({ error: "full_name is required" }, 400);

    const { data, error } = await sb
      .from("profiles")
      .update({ full_name: fullName, updated_at: new Date().toISOString() })
      .eq("id", user.id)
      .select("id, email, full_name, created_at, updated_at")
      .single();
    if (error) return json({ error: error.message }, 500);
    return json({ profile: data });
  }

  return json({ error: "Method not allowed" }, 405);
});
