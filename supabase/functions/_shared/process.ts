// Shared AP inbox processing: used by the `inbox` function (Process inbox button) and the `scheduler` function.
// Reads Outlook inbox mail, classifies it with Claude and forwards confident invoices to the Smart AP address.
// Each message is processed once per mailbox (unique mailbox_id + message_id). With reprocess, emails already
// classified are classified again unless they were forwarded or marked done, so nothing is forwarded twice.
import { encodeBase64 } from "jsr:@std/encoding@1/base64";
import { type AttachmentFile, classifyEmail, type Classification } from "./classify.ts";
import { logApi, logStep } from "./log.ts";
import { attachmentBytes, forwardMessage, type GraphMessage, graphToken, listAttachments, listInbox, sendMail } from "./microsoft.ts";
import { buildReport, type ReportItem } from "./report.ts";
import { adminClient } from "./supabase.ts";

// Keeps one run well inside the Edge Function time limit; the UI offers to run again when more remain.
const MAX_PER_RUN = 15;
const CONCURRENCY = 4;
export const STATUSES = ["processing", "action_needed", "done", "no_action", "error"];
const RECHECKABLE = ["no_action", "error", "action_needed"];

// Attachments sent to Claude. Small images are usually logos / signatures and are skipped.
const MAX_FILES = 4;
const MAX_FILE_BYTES = 10 * 1024 * 1024;
const MAX_TOTAL_BYTES = 20 * 1024 * 1024;
const MIN_IMAGE_BYTES = 15 * 1024;
const IMAGE_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp"];

// Every step of a run is logged (Logs > Steps) with the same run id, so one run can be followed end to end.
interface Run { id: string; userId: string; mailbox: string; trigger: string }
const step = (run: Run, location: string, message: string, context: Record<string, unknown> = {}) =>
  logStep(location, message, run.userId, { run: run.id, mailbox: run.mailbox, trigger: run.trigger, ...context });

// appUrl: where the app is hosted, for the "Open the Inbox" link in the summary email (null leaves it out).
export async function processMailbox(
  userId: string, mailboxId: string, hours: number, reprocess: boolean,
  trigger: "manual" | "schedule" = "manual", appUrl: string | null = null,
) {
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

  const run: Run = { id: crypto.randomUUID().slice(0, 8), userId, mailbox: mailbox.email_address, trigger };
  const started = Date.now();
  await step(run, "process/start",
    `Run started (${trigger === "schedule" ? "scheduled" : "manual"}) for ${mailbox.email_address}: last ${hours}h${reprocess ? ", re-checking processed emails" : ""}`);

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
  await step(run, "process/fetch",
    `Found ${messages.length} email${messages.length === 1 ? "" : "s"} in the inbox; ${fresh.length} to process` +
      (fresh.length > batch.length ? ` (${batch.length} now, ${fresh.length - batch.length} in the next run)` : ""),
    { skipped_already_processed: messages.length - fresh.length });

  const counts: Record<string, number> = {};
  const items: ReportItem[] = [];
  const queue = [...batch];
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, queue.length) }, async () => {
    for (let msg = queue.shift(); msg; msg = queue.shift()) {
      const item = await processMessage(db, run, mailbox, token, msg, seenById.get(msg.id));
      if (!item) continue;
      items.push(item);
      counts[item.status] = (counts[item.status] ?? 0) + 1;
    }
  }));

  await db.from("mailboxes").update({ last_processed_at: new Date().toISOString() }).eq("id", mailbox.id);
  const remaining = fresh.length - batch.length;
  const report = items.length ? await sendReport(run, mailbox, token, items, remaining, appUrl) : "skipped";
  const result = {
    mailbox: mailbox.email_address,
    scanned: messages.length,
    processed: batch.length,
    remaining,
    counts,
    report,
  };
  const summary = Object.entries(counts).map(([k, n]) => `${n} ${k.replace("_", " ")}`).join(", ") || "nothing new";
  await step(run, "process/finish", `Run finished in ${Math.round((Date.now() - started) / 1000)}s: ${summary}`, result);
  return result;
}

// Emails the run summary from the mailbox to itself (the mailbox owner). Needs Mail.Send, like forwarding.
async function sendReport(
  run: Run, mailbox: Mailbox, token: string, items: ReportItem[], remaining: number, appUrl: string | null,
): Promise<"sent" | "failed"> {
  const order = ["action_needed", "error", "done", "no_action"];
  items.sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status));
  const { subject, html } = buildReport(mailbox.email_address, items, remaining, run.trigger, appUrl);
  try {
    await sendMail(token, mailbox.email_address, subject, html);
    await step(run, "process/report", `Summary email sent to ${mailbox.email_address}: "${subject}"`);
    return "sent";
  } catch (err) {
    await logApi("process/report", err, { userId: run.userId, level: "warn", context: { run: run.id, mailbox: mailbox.email_address } });
    return "failed";
  }
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

// Returns the outcome for the summary email, or null if another run already claimed the message.
async function processMessage(
  db: Db, run: Run, mailbox: Mailbox, token: string, msg: GraphMessage, existing?: ExistingRow,
): Promise<ReportItem | null> {
  const userId = run.userId;
  const subject = msg.subject || "(no subject)";
  const rowId = await claim(db, userId, mailbox, msg, existing);
  if (!rowId) {
    await step(run, "process/email", `Skipped "${subject}": already being processed by another run`);
    return null;
  }
  const from = msg.from?.emailAddress ?? {};
  const email = { email_id: rowId, subject, from: from.address };
  const item: ReportItem = {
    subject,
    from: from.name ? `${from.name} <${from.address}>` : from.address ?? "unknown",
    category: null,
    summary: null,
    next_step: null,
    action: null,
    status: "error",
    web_link: msg.webLink ?? null,
  };

  try {
    const attachments = await gatherAttachments(token, msg);
    await step(run, "process/email",
      `Reading "${subject}" from ${from.address ?? "unknown"}` +
        (attachments.names.length
          ? `: ${attachments.names.length} attachment${attachments.names.length === 1 ? "" : "s"}, ${attachments.files.length} sent to Claude`
          : ": no attachments"),
      { ...email, attachments: attachments.names, sent_to_claude: attachments.files.map((f) => f.name), too_large: attachments.skipped });

    const c = await classifyEmail({
      from: from.name ? `${from.name} <${from.address}>` : from.address ?? "unknown",
      subject: msg.subject ?? "",
      received: msg.receivedDateTime,
      attachmentNames: attachments.names,
      files: attachments.files,
      skippedFiles: attachments.skipped,
      body: msg.body?.content ?? "",
    });
    await step(run, "process/classify", `Classified "${subject}" as ${c.category} (${c.confidence} confidence): ${c.summary}`,
      { ...email, category: c.category, confidence: c.confidence, vendor: c.vendor, amount: c.amount, next_step: c.next_step });

    const { status, action } = await takeAction(run, mailbox, token, msg, c);
    await step(run, "process/action", `"${subject}" → ${status.replace("_", " ")}${action ? `: ${action}` : ""}`, { ...email, status });

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
    return { ...item, category: c.category, summary: c.summary, next_step: c.next_step, action, status };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await logApi("inbox/process", err, { userId, context: { run: run.id, mailbox: mailbox.email_address, ...email } });
    await db.from("processed_emails").update({ status: "error", error: message, updated_at: new Date().toISOString() }).eq("id", rowId);
    return { ...item, status: "error", action: message };
  }
}

// The only automatic action: forward confident invoices to the mailbox's configured Smart AP address.
// The recipient always comes from settings, never from the email or the model.
async function takeAction(run: Run, mailbox: Mailbox, token: string, msg: GraphMessage, c: Classification) {
  const userId = run.userId;
  if (c.category === "other") return { status: "no_action", action: null };
  if (c.category !== "invoice") return { status: "action_needed", action: null };

  if (c.confidence === "low") return { status: "action_needed", action: "Not forwarded: unsure this is an invoice. Check it, then forward to AP." };
  if (!mailbox.smart_ap_address) return { status: "action_needed", action: "Not forwarded: no Smart AP address set for this mailbox." };

  try {
    const details = [c.vendor, c.invoice_number && `#${c.invoice_number}`, c.amount, c.due_date && `due ${c.due_date}`].filter(Boolean).join(", ");
    await forwardMessage(token, msg.id, mailbox.smart_ap_address, `Forwarded automatically by SummarizeMail: invoice${details ? ` (${details})` : ""}.`);
    return { status: "done", action: `Forwarded to ${mailbox.smart_ap_address}` };
  } catch (err) {
    await logApi("inbox/forward", err, { userId, level: "warn", context: { run: run.id, mailbox: mailbox.email_address, subject: msg.subject } });
    const hint = String(err).includes("403") ? " Reconnect the mailbox to allow sending." : "";
    return { status: "action_needed", action: `Forward failed; forward it manually.${hint}` };
  }
}
