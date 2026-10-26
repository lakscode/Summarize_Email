// GET    ?source=ui|api&kind=problems|steps&run=<id>&limit=200 -> { logs: [...] } (caller's own logs only, via RLS)
//        newest first; a single run (?run=) is returned oldest first so its steps read in order
// POST   { level, message, location, context } -> record a UI error for the caller
// DELETE -> clear the caller's logs
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { callerUser } from "../_shared/supabase.ts";

const LEVELS = ["error", "warn", "info"];

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;
  const url = new URL(req.url);

  try {
    if (req.method === "GET") {
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit")) || 200, 1), 1000);
      const run = url.searchParams.get("run");
      let query = sb
        .from("app_logs")
        .select("id, source, level, location, message, context, created_at")
        .order("created_at", { ascending: !!run })
        .order("id", { ascending: !!run })
        .limit(limit);
      const source = url.searchParams.get("source");
      if (source === "ui" || source === "api") query = query.eq("source", source);
      const kind = url.searchParams.get("kind");
      if (kind === "problems") query = query.in("level", ["error", "warn"]);
      if (kind === "steps") query = query.eq("level", "info");
      if (run) query = query.eq("context->>run", run);
      const { data, error } = await query;
      if (error) return await fail("logs/list", error.message, 500, { userId: user.id });
      return json({ logs: data });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      const message = String(body.message ?? "").trim();
      if (!message) return json({ error: "message is required" }, 400);
      const context = body.context && typeof body.context === "object" ? body.context : {};
      const { error } = await sb.from("app_logs").insert({
        user_id: user.id,
        source: "ui",
        level: LEVELS.includes(body.level) ? body.level : "error",
        location: String(body.location ?? "").slice(0, 200) || null,
        message: message.slice(0, 2000),
        // Cap stored context so a runaway client can't write huge rows.
        context: JSON.stringify(context).length > 8000 ? { truncated: true } : context,
      });
      if (error) return await fail("logs/write", error.message, 500, { userId: user.id });
      return json({ ok: true }, 201);
    }

    if (req.method === "DELETE") {
      const { error } = await sb.from("app_logs").delete().eq("user_id", user.id);
      if (error) return await fail("logs/clear", error.message, 500, { userId: user.id });
      return json({ cleared: true });
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("logs", err, { userId: user.id });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
