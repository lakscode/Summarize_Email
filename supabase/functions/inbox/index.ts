// AP inbox processing for the signed-in user's connected mailboxes.
//   GET   ?status=&mailbox_id=                          -> { emails: [...] } newest first
//   POST  {action: "process", mailbox_id, lookback_hours, reprocess?} -> read inbox mail, classify with Claude,
//                                                          forward invoices to the Smart AP address
//   PATCH {id, status: "done" | "action_needed"}       -> employee marks an item handled / reopens it
// Each message is processed once per mailbox (unique mailbox_id + message_id). With reprocess, emails already
// classified are classified again unless they were forwarded or marked done, so nothing is forwarded twice.
import { encodeBase64 } from "jsr:@std/encoding@1/base64";
import { type AttachmentFile, classifyEmail, type Classification } from "../_shared/classify.ts";
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { attachmentBytes, forwardMessage, type GraphMessage, graphToken, listAttachments, listInbox } from "../_shared/microsoft.ts";
import { adminClient, callerUser } from "../_shared/supabase.ts";

// Keeps one run well inside the Edge Function time limit; the UI offers to run again when more remain.
const MAX_PER_RUN = 15;
const CONCURRENCY = 4;
const STATUSES = ["processing", "action_needed", "done", "no_action", "error"];
const RECHECKABLE = ["no_action", "error", "action_needed"];

// Attachments sent to Claude. Small images are usually logos / signatures and are skipped.
const MAX_FILES = 4;
const MAX_FILE_BYTES = 10 * 1024 * 1024;
const MAX_TOTAL_BYTES = 20 * 1024 * 1024;
const MIN_IMAGE_BYTES = 15 * 1024;
const IMAGE_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp"];

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
      return json(await processMailbox(user.id, body.mailbox_id, hours, body.reprocess === true));
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("inbox", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

async function processMailbox(userId: string, mailboxId: string, hours: number, reprocess: boolean) {
  const db = adminClient();
  await db.rpc("expire_mailbox_access");
  const { data: mailbox, error } = await db.from("mailboxes")
    .select("id, email_address, credentials_secret_id, smart_ap_address, access_expires_at")
    .eq("id", mailboxId).eq("owner_user_id", userId).maybeSingle();
  if (error) throw new Error(error.message);
  if (!mailbox) throw new Error("Mailbox not found");
  if (!mailbox.credentials_secret_id) {
    throw new Error(`Access to ${mailbox.email_address} has expired. Reconnect it on the dashboard.`);
  }

  const token = await graphToken(db, mailbox);
  const messages = await listInbox(token, new Date(Date.now() - hours * 3600_000));

  const { data: seen, error: seenError } = messages.length
    ? await db.from("processed_emails").select("id, message_id, status, action_taken")
      .eq("mailbox_id", mailbox.id).in("message_id", messages.map((m) => m.id))
    : { data: [], error: null };
  if (seenError) throw new Error(seenError.message);
  const seenById = new Map<string, ExistingRow>((seen as ExistingRow[] ?? []).map((r) => [r.message_id, r]));
  const recheckable = (row: ExistingRow) =>
    RECHECKABLE.includes(row.status) && !(row.action_taken ?? "").startsWith("Forwarded");
  const fresh = messages.filter((m) => {
    const row = seenById.get(m.id);
    return !row || (reprocess && recheckable(row));
  });
  const batch = fresh.slice(0, MAX_PER_RUN);

  const counts: Record<string, number> = {};
  const queue = [...batch];
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, queue.length) }, async () => {
    for (let msg = queue.shift(); msg; msg = queue.shift()) {
      const status = await processMessage(db, userId, mailbox, token, msg, seenById.get(msg.id));
      if (status) counts[status] = (counts[status] ?? 0) + 1;
    }
  }));

  await db.from("mailboxes").update({ last_processed_at: new Date().toISOString() }).eq("id", mailbox.id);
  return {
    mailbox: mailbox.email_address,
    scanned: messages.length,
    processed: batch.length,
    remaining: fresh.length - batch.length,
    counts,
  };
}

type Db = ReturnType<typeof adminClient>;
interface Mailbox { id: string; email_address: string; credentials_secret_id: string | null; smart_ap_address: string | null }
interface ExistingRow { id: string; message_id: string; status: string; action_taken: string | null }

// Claims the message for this run: inserts a new row, or (re-check) flips an existing row back to processing
// only if it is still in the status we read, so two runs never work on the same email.
async function claim(db: Db, userId: string, mailbox: Mailbox, msg: GraphMessage, existing?: ExistingRow): Promise<string | null> {
  if (existing) {
    const { data, error } = await db.from("processed_emails")
      .update({ status: "processing", error: null, updated_at: new Date().toISOString() })
      .eq("id", existing.id).eq("status", existing.status)
      .select("id");
    if (error) throw new Error(error.message);
    return data?.[0]?.id ?? null;
  }
  const from = msg.from?.emailAddress ?? {};
  const { data, error } = await db.from("processed_emails").upsert({
    mailbox_id: mailbox.id,
    owner_user_id: userId,
    message_id: msg.id,
    received_at: msg.receivedDateTime,
    from_name: from.name ?? null,
    from_address: from.address?.toLowerCase() ?? null,
    subject: msg.subject ?? "",
    web_link: msg.webLink ?? null,
    status: "processing",
  }, { onConflict: "mailbox_id,message_id", ignoreDuplicates: true }).select("id");
  if (error) throw new Error(error.message);
  return data?.[0]?.id ?? null;
}

// All attachment names, plus the PDFs / images Claude should read (within size limits).
async function gatherAttachments(token: string, msg: GraphMessage) {
  if (!msg.hasAttachments) return { names: [] as string[], files: [] as AttachmentFile[], skipped: [] as string[] };
  const all = await listAttachments(token, msg.id);
  const files: AttachmentFile[] = [];
  const skipped: string[] = [];
  let total = 0;

  for (const a of all) {
    if (a["@odata.type"] && a["@odata.type"] !== "#microsoft.graph.fileAttachment") continue;
    const type = (a.contentType ?? "").toLowerCase();
    const isPdf = type === "application/pdf" || a.name.toLowerCase().endsWith(".pdf");
    const isImage = IMAGE_TYPES.includes(type) && !a.isInline && a.size >= MIN_IMAGE_BYTES;
    if (!isPdf && !isImage) continue;
    if (files.length >= MAX_FILES || a.size > MAX_FILE_BYTES || total + a.size > MAX_TOTAL_BYTES) {
      skipped.push(a.name);
      continue;
    }
    const bytes = await attachmentBytes(token, msg.id, a.id);
    total += bytes.length;
    files.push({
      name: a.name,
      mediaType: isPdf ? "application/pdf" : type as AttachmentFile["mediaType"],
      base64: encodeBase64(bytes),
    });
  }
  return { names: all.map((a) => a.name), files, skipped };
}

// Returns the final status, or null if another run already claimed the message.
async function processMessage(
  db: Db, userId: string, mailbox: Mailbox, token: string, msg: GraphMessage, existing?: ExistingRow,
): Promise<string | null> {
  const rowId = await claim(db, userId, mailbox, msg, existing);
  if (!rowId) return null;
  const from = msg.from?.emailAddress ?? {};

  try {
    const attachments = await gatherAttachments(token, msg);
    const c = await classifyEmail({
      from: from.name ? `${from.name} <${from.address}>` : from.address ?? "unknown",
      subject: msg.subject ?? "",
      received: msg.receivedDateTime,
      attachmentNames: attachments.names,
      files: attachments.files,
      skippedFiles: attachments.skipped,
      body: msg.body?.content ?? "",
    });
    const { status, action } = await takeAction(userId, mailbox, token, msg, c);

    const { error } = await db.from("processed_emails").update({
      attachment_names: attachments.names,
      category: c.category,
      confidence: c.confidence,
      vendor: c.vendor || null,
      invoice_number: c.invoice_number || null,
      amount: c.amount || null,
      due_date: c.due_date || null,
      summary: c.summary,
      next_step: c.next_step,
      action_taken: action,
      status,
      updated_at: new Date().toISOString(),
    }).eq("id", rowId);
    if (error) throw new Error(error.message);
    return status;
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await logApi("inbox/process", err, { userId, context: { mailbox: mailbox.email_address, subject: msg.subject } });
    await db.from("processed_emails").update({ status: "error", error: message, updated_at: new Date().toISOString() }).eq("id", rowId);
    return "error";
  }
}

// The only automatic action: forward confident invoices to the mailbox's configured Smart AP address.
// The recipient always comes from settings, never from the email or the model.
async function takeAction(userId: string, mailbox: Mailbox, token: string, msg: GraphMessage, c: Classification) {
  if (c.category === "other") return { status: "no_action", action: null };
  if (c.category !== "invoice") return { status: "action_needed", action: null };

  if (c.confidence === "low") return { status: "action_needed", action: "Not forwarded: unsure this is an invoice. Check it, then forward to AP." };
  if (!mailbox.smart_ap_address) return { status: "action_needed", action: "Not forwarded: no Smart AP address set for this mailbox." };

  try {
    const details = [c.vendor, c.invoice_number && `#${c.invoice_number}`, c.amount, c.due_date && `due ${c.due_date}`].filter(Boolean).join(", ");
    await forwardMessage(token, msg.id, mailbox.smart_ap_address, `Forwarded automatically by SummarizeMail: invoice${details ? ` (${details})` : ""}.`);
    return { status: "done", action: `Forwarded to ${mailbox.smart_ap_address}` };
  } catch (err) {
    await logApi("inbox/forward", err, { userId, level: "warn", context: { mailbox: mailbox.email_address, subject: msg.subject } });
    const hint = String(err).includes("403") ? " Reconnect the mailbox to allow sending." : "";
    return { status: "action_needed", action: `Forward failed; forward it manually.${hint}` };
  }
}
