// AP inbox processing for the signed-in user's connected mailboxes.
//   GET   ?status=&mailbox_id=                          -> { emails: [...] } newest first
//   GET   ?preview=<email id>                           -> { email: {...} } the message read live from the mailbox
//   GET   ?history=<email id>                           -> { history: [...] } who processed it and when, newest first
//   POST  {action: "process", mailbox_id, lookback_hours, reprocess?, all?, started_at?} -> process a batch now
//         (see _shared/process.ts); with all, the page calls again until remaining is 0
//   POST  {action: "preview", mailbox_id, lookback_hours} -> { new, already_processed, ... } to confirm a re-check first
//   PATCH {id, status: "done" | "action_needed"}       -> employee marks an item handled / reopens it
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi, logStep } from "../_shared/log.ts";
import { previewEmail, previewMailbox, processMailbox, processorOf, STATUSES } from "../_shared/process.ts";
import { isNotFound } from "../_shared/mail.ts";
import { adminClient, callerUser } from "../_shared/supabase.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;
  const url = new URL(req.url);

  try {
    if (req.method === "GET" && url.searchParams.has("preview")) {
      const id = url.searchParams.get("preview")!;
      try {
        const email = await previewEmail(user.id, id);
        return email ? json({ email }) : json({ error: "Email not found" }, 404);
      } catch (err) {
        // Gone from the inbox: deleted, or moved (Outlook gives a moved message a new id).
        if (isNotFound(err)) {
          return await fail("inbox/preview", "This email is no longer in the inbox. It may have been moved or deleted.", 404, { userId: user.id, context: { email_id: id } });
        }
        throw err;
      }
    }

    if (req.method === "GET" && url.searchParams.has("history")) {
      const { data, error } = await sb.from("email_processing_log")
        .select("processed_at, processed_by_email, processed_by_name, trigger, run_id, status, category, action, error")
        .eq("processed_email_id", url.searchParams.get("history"))
        .order("processed_at", { ascending: false })
        .limit(100);
      if (error) return await fail("inbox/history", error.message, 500, { userId: user.id });
      return json({ history: data });
    }

    if (req.method === "GET") {
      let query = sb
        .from("processed_emails")
        .select("id, mailbox_id, received_at, from_name, from_address, subject, web_link, attachment_names, category, confidence, vendor, invoice_number, amount, due_date, summary, next_step, action_taken, status, error, processed_by_email, processed_by_name, processed_at")
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
        .select("id, status, mailbox_id, owner_user_id, subject, mailboxes(email_address)")
        .maybeSingle();
      if (error) return await fail("inbox/update", error.message, 500, { userId: user.id, context: { email_id: body.id } });
      if (!data) return json({ error: "Email not found" }, 404);
      // Marking done / reopening counts as processing it: this user and time become the email's "last processed".
      const db = adminClient();
      const processor = await processorOf(db, user.id);
      const by = { processed_by: user.id, processed_by_email: processor.email ?? user.email ?? null, processed_by_name: processor.name ?? user.email ?? null };
      const processedAt = new Date().toISOString();
      const { error: byError } = await db.from("processed_emails").update({ ...by, processed_at: processedAt }).eq("id", data.id);
      const mailboxAddress = (data.mailboxes as { email_address?: string } | null)?.email_address ?? null;
      const logContext = { email_id: data.id, subject: data.subject, mailbox: mailboxAddress };
      if (byError) await logApi("inbox/update", byError.message, { userId: user.id, level: "warn", context: logContext });
      const { error: logError } = await db.from("email_processing_log").insert({
        processed_email_id: data.id,
        mailbox_id: data.mailbox_id,
        mailbox_address: mailboxAddress,
        subject: data.subject,
        owner_user_id: data.owner_user_id,
        ...by,
        processed_at: processedAt,
        trigger: "user",
        status: data.status,
        action: data.status === "done" ? "Marked done" : "Reopened",
      });
      if (logError) await logApi("inbox/update", logError.message, { userId: user.id, level: "warn", context: logContext });
      await logStep("inbox/update", `"${data.subject || "(no subject)"}" ${data.status === "done" ? "marked done" : "reopened"} by ${by.processed_by_name ?? "unknown"}`,
        user.id, { ...logContext, status: data.status });
      return json({ email: { id: data.id, status: data.status } });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (!["process", "preview"].includes(body.action) || !body.mailbox_id) {
        return json({ error: "action ('process' | 'preview') and mailbox_id are required" }, 400);
      }
      const hours = Math.min(Math.max(Number(body.lookback_hours) || 24, 1), 24 * 14);
      if (body.action === "preview") return json(await previewMailbox(user.id, body.mailbox_id, hours));
      const origin = req.headers.get("Origin");
      const appUrl = origin && /^https?:\/\/[^/]+$/.test(origin) ? origin : null;
      // all: the page keeps calling until nothing remains, passing back the first call's started_at.
      const startedAt = typeof body.started_at === "string" && !isNaN(Date.parse(body.started_at)) ? body.started_at : null;
      const untilDone = body.all === true ? { startedAt } : null;
      return json(await processMailbox(user.id, body.mailbox_id, hours, body.reprocess === true, "manual", appUrl, untilDone));
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("inbox", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
