// GET    ?source=ui|api&kind=problems|steps|emails&run=<id>&email=<id>&mailbox=&user=&from=&to=&limit=200 -> { logs }
//        newest first; a single run (?run=) or email (?email=, a processed email id) is returned oldest first so
//        its steps read in order.
//        kind=emails: the email processing history (email_processing_log) instead of app_logs.
//        mailbox: an address; user: a user id; from / to: ISO times (from inclusive, to exclusive).
// GET    ?filters=1 -> { users, mailboxes, others } the choices for the filters (mailboxes: the caller's own)
//        Whose logs: everyone's for a super admin, the organization's members' for an organization admin, otherwise
//        only the caller's own (others: false).
// POST   { level, message, location, context } -> record a UI error for the caller
// DELETE -> clear the caller's logs
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { adminClient, callerUser } from "../_shared/supabase.ts";

const LEVELS = ["error", "warn", "info"];
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const time = (v: string | null) => v && !isNaN(Date.parse(v)) ? new Date(v).toISOString() : null;

// The users whose logs the caller may see; null: everyone (super admin).
async function visibleUsers(userId: string): Promise<string[] | null> {
  const db = adminClient();
  const { data: superAdmin } = await db.from("super_admins").select("user_id").eq("user_id", userId).maybeSingle();
  if (superAdmin) return null;
  const { data: own } = await db.from("organization_members").select("org_id, role").eq("user_id", userId).maybeSingle();
  if (own?.role !== "admin") return [userId];
  const { data, error } = await db.from("organization_members").select("user_id").eq("org_id", own.org_id);
  if (error) throw new Error(error.message);
  return data.map((r) => r.user_id);
}

// Display names (profile name, else email) of these users.
async function names(ids: string[]): Promise<Map<string, { name: string; email: string | null }>> {
  const unique = [...new Set(ids.filter(Boolean))];
  if (!unique.length) return new Map();
  const { data, error } = await adminClient().from("profiles").select("id, email, full_name").in("id", unique);
  if (error) throw new Error(error.message);
  return new Map(data.map((p) => [p.id, { name: p.full_name?.trim() || p.email || "Unknown user", email: p.email }]));
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;
  const url = new URL(req.url);

  try {
    if (req.method === "GET" && url.searchParams.has("filters")) {
      const visible = await visibleUsers(user.id);
      const db = adminClient();
      let users = db.from("profiles").select("id").order("email").limit(1000);
      // Mailboxes: only the caller's own (connected under their account, shared ones included), also for admins.
      const boxes = db.from("mailboxes").select("email_address").eq("owner_user_id", user.id).order("email_address");
      if (visible) users = users.in("id", visible);
      const [u, b] = await Promise.all([users, boxes]);
      if (u.error || b.error) return await fail("logs/filters", (u.error ?? b.error)!.message, 500, { userId: user.id });
      const byId = await names(u.data.map((r) => r.id));
      return json({
        others: visible === null || visible.length > 1,
        users: [...byId].map(([id, p]) => ({ id, ...p })).sort((x, y) => x.name.localeCompare(y.name)),
        mailboxes: [...new Set(b.data.map((r) => r.email_address))],
      });
    }

    if (req.method === "GET") {
      const limit = Math.min(Math.max(Number(url.searchParams.get("limit")) || 200, 1), 1000);
      const run = url.searchParams.get("run");
      const emailId = url.searchParams.get("email");
      if (emailId && !UUID.test(emailId)) return json({ error: "email must be an email id" }, 400);
      const ordered = !!run || !!emailId;
      const kind = url.searchParams.get("kind");
      const mailbox = url.searchParams.get("mailbox") || null;
      const userFilter = url.searchParams.get("user");
      if (userFilter && !UUID.test(userFilter)) return json({ error: "user must be a user id" }, 400);
      const from = time(url.searchParams.get("from"));
      const to = time(url.searchParams.get("to"));

      const visible = await visibleUsers(user.id);
      if (userFilter && visible && !visible.includes(userFilter)) return json({ error: "You can't see that user's logs" }, 403);
      // Only the caller: their own client, so RLS applies as before. Others: the service role, scoped here.
      const own = visible?.length === 1 && visible[0] === user.id;
      const db = own ? sb : adminClient();

      if (kind === "emails") {
        let query = db.from("email_processing_log")
          .select("id, processed_at, mailbox_address, subject, processed_email_id, owner_user_id, processed_by, processed_by_name, processed_by_email, trigger, run_id, status, category, action, error")
          .order("processed_at", { ascending: ordered })
          .limit(limit);
        if (emailId) query = query.eq("processed_email_id", emailId);
        if (visible) query = query.in("owner_user_id", visible);
        if (userFilter) query = query.eq("processed_by", userFilter);
        if (mailbox) query = query.eq("mailbox_address", mailbox);
        if (run) query = query.eq("run_id", run);
        if (from) query = query.gte("processed_at", from);
        if (to) query = query.lt("processed_at", to);
        const { data, error } = await query;
        if (error) return await fail("logs/emails", error.message, 500, { userId: user.id });
        return json({ logs: data });
      }

      let query = db
        .from("app_logs")
        .select("id, user_id, source, level, location, message, context, created_at")
        .order("created_at", { ascending: ordered })
        .order("id", { ascending: ordered })
        .limit(limit);
      if (emailId) query = query.eq("context->>email_id", emailId);
      const source = url.searchParams.get("source");
      if (source === "ui" || source === "api") query = query.eq("source", source);
      if (kind === "problems") query = query.in("level", ["error", "warn"]);
      if (kind === "steps") query = query.eq("level", "info");
      if (run) query = query.eq("context->>run", run);
      if (visible) query = query.in("user_id", visible);
      if (userFilter) query = query.eq("user_id", userFilter);
      if (mailbox) query = query.eq("context->>mailbox", mailbox);
      if (from) query = query.gte("created_at", from);
      if (to) query = query.lt("created_at", to);
      const { data, error } = await query;
      if (error) return await fail("logs/list", error.message, 500, { userId: user.id });
      const byId = await names(data.map((r) => r.user_id));
      return json({ logs: data.map((r) => ({ ...r, user_name: byId.get(r.user_id)?.name ?? null })) });
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
