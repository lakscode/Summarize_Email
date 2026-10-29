// AP inbox processing for the signed-in user's connected mailboxes.
//   GET   ?status=&mailbox_id=                          -> { emails: [...] } newest first
//   POST  {action: "process", mailbox_id, lookback_hours, reprocess?} -> process now (see _shared/process.ts)
//   PATCH {id, status: "done" | "action_needed"}       -> employee marks an item handled / reopens it
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { processMailbox, STATUSES } from "../_shared/process.ts";
import { callerUser } from "../_shared/supabase.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;
  const url = new URL(req.url);

  try {
    if (req.method === "GET") {
      let query = sb
        .from("processed_emails")
        .select("id, mailbox_id, received_at, from_name, from_address, subject, web_link, attachment_names, category, confidence, vendor, invoice_number, amount, due_date, summary, next_step, action_taken, status, error")
        .order("received_at", { ascending: false })
        .limit(300);
      const status = url.searchParams.get("status");
      if (status && STATUSES.includes(status)) query = query.eq("status", status);
      const mailboxId = url.searchParams.get("mailbox_id");
      if (mailboxId) query = query.eq("mailbox_id", mailboxId);
      const { data, error } = await query;
      if (error) return await fail("inbox/list", error.message, 500, { userId: user.id });
      return json({ emails: data });
    }

    if (req.method === "PATCH") {
      const body = await req.json().catch(() => ({}));
      if (!body.id || !["done", "action_needed"].includes(body.status)) {
        return json({ error: "id and status (done | action_needed) are required" }, 400);
      }
      const { data, error } = await sb.from("processed_emails")
        .update({ status: body.status, updated_at: new Date().toISOString() })
        .eq("id", body.id)
        .select("id, status")
        .maybeSingle();
      if (error) return await fail("inbox/update", error.message, 500, { userId: user.id, context: { id: body.id } });
      if (!data) return json({ error: "Email not found" }, 404);
      return json({ email: data });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action !== "process" || !body.mailbox_id) return json({ error: "action 'process' and mailbox_id are required" }, 400);
      const hours = Math.min(Math.max(Number(body.lookback_hours) || 24, 1), 24 * 14);
      const origin = req.headers.get("Origin");
      const appUrl = origin && /^https?:\/\/[^/]+$/.test(origin) ? origin : null;
      return json(await processMailbox(user.id, body.mailbox_id, hours, body.reprocess === true, "manual", appUrl));
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("inbox", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
