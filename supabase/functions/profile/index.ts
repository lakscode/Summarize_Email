// GET  /functions/v1/profile -> the signed-in user's profile
// PATCH /functions/v1/profile { full_name } -> update it
// Queries run as the caller (their JWT is forwarded), so RLS on public.profiles applies.
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { callerUser } from "../_shared/supabase.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;

  try {
    if (req.method === "GET") {
      const { data, error } = await sb
        .from("profiles")
        .select("id, email, full_name, created_at, updated_at")
        .eq("id", user.id)
        .maybeSingle();
      if (error) return await fail("profile/get", error.message, 500, { userId: user.id });
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
      if (error) return await fail("profile/update", error.message, 500, { userId: user.id });
      return json({ profile: data });
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("profile", err, { userId: user.id });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
