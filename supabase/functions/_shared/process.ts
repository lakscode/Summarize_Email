// Shared AP / AR inbox processing: used by the `inbox` function (Process inbox button) and the `scheduler` function.
// Reads Outlook or Gmail inbox mail (see mail.ts), classifies it with Claude for the mailbox's type (AP or AR) and
// forwards confident emails in the mailbox's forward_categories to its destination address (smart_ap_address).
// Each message is processed once per mailbox (unique mailbox_id + message_id). With reprocess, emails already
// classified are classified again unless they were forwarded or marked done, so nothing is forwarded twice.
import { encodeBase64 } from "./base64.ts";
import { APP_NAME } from "./brand.ts";
import PostalMime from "npm:postal-mime@3.0.1";
import { type AttachmentFile, type AttachmentText, classifyEmail, type Classification, type MailboxType } from "./classify.ts";
import { logApi, logStep } from "./log.ts";
import {
  embedInlineImages, htmlToText, isEmailAttachment, type MailClient, mailClient, type MailMessage, type Provider,
} from "./mail.ts";
import { buildReport, type ReportItem } from "./report.ts";
import { adminClient } from "./supabase.ts";

// Keeps one call well inside the Edge Function time limit; the UI calls again until none remain (see `untilDone`).
const MAX_PER_RUN = 15;
// Most inbox emails one look-back period covers.
const MAX_INBOX = 500;
const CONCURRENCY = 4;
export const STATUSES = ["processing", "action_needed", "done", "no_action", "error"];
const RECHECKABLE = ["no_action", "error", "action_needed"];
// A claim older than this is from a run that was cut off (e.g. the Edge Function hit its time limit): the email
// counts as not processed and the next run takes it over. A batch never takes this long.
const STALE_CLAIM_MS = 15 * 60_000;

// Attachments sent to Claude. Small images are usually logos / signatures and are skipped.
const MAX_FILES = 4;
const MAX_FILE_BYTES = 10 * 1024 * 1024;
const MAX_TOTAL_BYTES = 20 * 1024 * 1024;
const MIN_IMAGE_BYTES = 15 * 1024;
const IMAGE_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp"];
// Attached emails (.eml / Outlook items) and text attachments are read as text, within these limits.
const MAX_TEXTS = 6;
const MAX_TEXT_CHARS = 30_000;
const MAX_EMAIL_BYTES = 25 * 1024 * 1024;
const MAX_EMAIL_DEPTH = 2; // an attached email inside an attached email is still read; deeper ones are not
const TEXT_TYPES = ["text/plain", "text/csv", "text/tab-separated-values", "text/html", "text/xml", "application/xml", "application/json"];
const TEXT_EXTENSION = /\.(txt|csv|tsv|html?|xml|json)$/i;

// Every step of a run is logged (Logs > Steps) with the same run id, so one run can be followed end to end.
// userEmail / userName: who processed, recorded on each email, its mailbox and in email_processing_log (the owner,
// also for scheduled runs).
interface Run {
  id: string; userId: string; userEmail: string | null; userName: string | null; mailbox: string; trigger: "manual" | "schedule";
}

// A user's email and display name (profile name, else email) for the "processed by" fields.
export async function processorOf(db: Db, userId: string) {
  const { data } = await db.from("profiles").select("email, full_name").eq("id", userId).maybeSingle();
  const email = data?.email ?? null;
  return { email, name: data?.full_name?.trim() || email };
}
const step = (run: Run, location: string, message: string, context: Record<string, unknown> = {}) =>
  logStep(location, message, run.userId, { run: run.id, mailbox: run.mailbox, trigger: run.trigger, ...context });

async function loadMailbox(db: Db, userId: string, mailboxId: string): Promise<Mailbox> {
  await db.rpc("expire_mailbox_access");
  const { data: mailbox, error } = await db.from("mailboxes")
    .select("id, provider, auth_type, connection_id, email_address, delegate_address, credentials_secret_id, smart_ap_address, access_expires_at, mailbox_type, forward_categories, send_summary, summary_address")
    .eq("id", mailboxId).eq("owner_user_id", userId).maybeSingle();
  if (error) throw new Error(error.message);
  if (!mailbox) throw new Error("Mailbox not found");
  if (!mailbox.credentials_secret_id && !mailbox.connection_id) {
    throw new Error(`Access to ${mailbox.email_address} has expired. Reconnect it on the dashboard.`);
  }
  return mailbox;
}

// The ids of the inbox messages of the last `hours`, with the processed_emails row of those already processed.
async function listInbox(db: Db, mailbox: Mailbox, client: MailClient, hours: number) {
  const ids = await client.listInboxIds(new Date(Date.now() - hours * 3600_000), MAX_INBOX);
  const seen: ExistingRow[] = [];
  for (let i = 0; i < ids.length; i += 100) { // keeps the request URL short
    const { data, error } = await db.from("processed_emails").select("id, message_id, status, action_taken, processed_at, updated_at")
      .eq("mailbox_id", mailbox.id).in("message_id", ids.slice(i, i + 100));
    if (error) throw new Error(error.message);
    seen.push(...data as ExistingRow[]);
  }
  return { ids, seenById: new Map(seen.map((r) => [r.message_id, r])) };
}

// Already processed emails that a re-check classifies again: never ones that were forwarded or marked done.
const recheckable = (row: ExistingRow) =>
  RECHECKABLE.includes(row.status) && !(row.action_taken ?? "").startsWith("Forwarded");
const staleClaim = (row: ExistingRow) =>
  row.status === "processing" && Date.now() - new Date(row.updated_at).getTime() > STALE_CLAIM_MS;

// What a run of the last `hours` would pick up, so the UI can ask before processing emails again.
export async function previewMailbox(userId: string, mailboxId: string, hours: number) {
  const db = adminClient();
  const mailbox = await loadMailbox(db, userId, mailboxId);
  const { ids, seenById } = await listInbox(db, mailbox, await mailClient(db, mailbox), hours);
  const rows = [...seenById.values()].filter((r) => !staleClaim(r));
  return {
    mailbox: mailbox.email_address,
    scanned: ids.length,
    new: ids.length - rows.length,
    already_processed: rows.filter(recheckable).length,
    handled: rows.filter((r) => !recheckable(r)).length,
  };
}

// A processed email as the Inbox preview shows it, read live from the mailbox (null if it isn't the user's).
export async function previewEmail(userId: string, emailId: string) {
  const db = adminClient();
  const { data: row, error } = await db.from("processed_emails").select("mailbox_id, message_id, web_link")
    .eq("id", emailId).eq("owner_user_id", userId).maybeSingle();
  if (error) throw new Error(error.message);
  if (!row) return null;
  const mailbox = await loadMailbox(db, userId, row.mailbox_id);
  const m = await (await mailClient(db, mailbox)).getMessage(row.message_id);
  return {
    subject: m.subject ?? "",
    from: m.from ?? "",
    to: m.to ?? "",
    cc: m.cc ?? "",
    received: m.received,
    html: m.html === null ? null : embedInlineImages(m.html, m.inlineImages),
    text: m.text,
    attachments: m.attachments.map(({ name, size, contentType }) => ({ name, size, content_type: contentType ?? null })),
    web_link: m.webLink ?? row.web_link,
  };
}

// appUrl: where the app is hosted, for the "Open the Inbox" link in the summary email (null leaves it out).
// untilDone: the caller (the Inbox page) calls again while `remaining` > 0, passing back the `started_at` of the first
// call. Emails processed since then aren't re-checked again, and the one summary email goes out with the last call.
export async function processMailbox(
  userId: string, mailboxId: string, hours: number, reprocess: boolean,
  trigger: "manual" | "schedule" = "manual", appUrl: string | null = null,
  untilDone: { startedAt: string | null } | null = null,
) {
  const db = adminClient();
  const mailbox = await loadMailbox(db, userId, mailboxId);

  const processor = await processorOf(db, userId);
  const run: Run = {
    id: crypto.randomUUID().slice(0, 8), userId, userEmail: processor.email, userName: processor.name,
    mailbox: mailbox.email_address, trigger,
  };
  const started = Date.now();
  const startedAt = untilDone?.startedAt ?? new Date(started).toISOString();
  await step(run, "process/start",
    `Run started (${trigger === "schedule" ? "scheduled" : "manual"}) for ${mailbox.email_address}: last ${hours}h${reprocess ? ", re-checking processed emails" : ""}`);

  const client = await mailClient(db, mailbox);
  const { ids, seenById } = await listInbox(db, mailbox, client, hours);
  const doneThisRun = (row: ExistingRow) => !!row.processed_at && new Date(row.processed_at) >= new Date(startedAt);
  const fresh = ids.filter((id) => {
    const row = seenById.get(id);
    return !row || staleClaim(row) || (reprocess && recheckable(row) && !doneThisRun(row));
  });
  const batch = await client.getMessages(fresh.slice(0, MAX_PER_RUN));
  const remaining = Math.max(fresh.length - MAX_PER_RUN, 0);
  await step(run, "process/fetch",
    `Found ${ids.length} email${ids.length === 1 ? "" : "s"} in the inbox; ${fresh.length} to process` +
      (remaining ? ` (${batch.length} now, ${remaining} in the next batch)` : ""),
    { skipped_already_processed: ids.length - fresh.length });

  const counts: Record<string, number> = {};
  const items: ReportItem[] = [];
  const queue = [...batch];
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, queue.length) }, async () => {
    for (let msg = queue.shift(); msg; msg = queue.shift()) {
      const item = await processMessage(db, run, mailbox, client, msg, seenById.get(msg.id));
      if (!item) continue;
      items.push(item);
      counts[item.status] = (counts[item.status] ?? 0) + 1;
    }
  }));

  await db.from("mailboxes").update({
    last_processed_at: new Date().toISOString(), last_processed_by: userId, last_processed_by_name: run.userName,
  }).eq("id", mailbox.id);
  // Only when the mailbox's "Send summary email" setting is on, it has a recipient, and there is something to report.
  // Run until done: only with the last batch, covering every batch of the run.
  const pending = !!untilDone && remaining > 0;
  const reportItems = untilDone && !pending ? await itemsSince(db, mailbox.id, startedAt) : items;
  const to = mailbox.send_summary && !pending ? await summaryRecipient(db, userId, mailbox) : null;
  const report = !mailbox.send_summary ? "off" : pending ? "pending" : !to ? "no_recipient"
    : reportItems.length ? await sendReport(run, mailbox, client, to, reportItems, remaining, appUrl) : "skipped";
  const result = {
    mailbox: mailbox.email_address,
    scanned: ids.length,
    processed: batch.length,
    remaining,
    counts,
    report,
    started_at: startedAt,
  };
  const summary = Object.entries(counts).map(([k, n]) => `${n} ${k.replace("_", " ")}`).join(", ") || "nothing new";
  await step(run, "process/finish", `Run finished in ${Math.round((Date.now() - started) / 1000)}s: ${summary}`, result);
  return result;
}

// The emails of a run-until-done (processed since its first batch started), for its summary email.
async function itemsSince(db: Db, mailboxId: string, since: string): Promise<ReportItem[]> {
  const { data, error } = await db.from("processed_emails")
    .select("subject, from_name, from_address, category, summary, next_step, action_taken, status, error, web_link")
    .eq("mailbox_id", mailboxId).gte("processed_at", since).neq("status", "processing")
    .order("received_at", { ascending: false }).limit(MAX_INBOX);
  if (error) throw new Error(error.message);
  return (data ?? []).map((r) => ({
    subject: r.subject || "(no subject)",
    from: r.from_name ? `${r.from_name} <${r.from_address}>` : r.from_address ?? "unknown",
    category: r.category,
    summary: r.summary,
    next_step: r.next_step,
    action: r.status === "error" ? r.error : r.action_taken,
    status: r.status,
    web_link: r.web_link,
  }));
}

// Where the summary goes: the mailbox's summary address if set; otherwise an own mailbox gets it itself, and a shared
// mailbox (read through another account or an organization connection) goes to the owner's account email.
async function summaryRecipient(db: Db, userId: string, mailbox: Mailbox): Promise<string | null> {
  if (mailbox.summary_address) return mailbox.summary_address;
  const shared = !!mailbox.delegate_address || mailbox.auth_type === "app";
  if (!shared) return mailbox.email_address;
  const { data, error } = await db.auth.admin.getUserById(userId);
  if (error) await logApi("process/report", error.message, { userId, level: "warn", context: { mailbox: mailbox.email_address } });
  return data?.user?.email?.toLowerCase() || null;
}

// Emails the run summary to `to`. Needs Mail.Send, like forwarding. For a shared mailbox connected by sign-in it is
// sent from the signed-in user's own mailbox.
async function sendReport(
  run: Run, mailbox: Mailbox, client: MailClient, to: string, items: ReportItem[], remaining: number, appUrl: string | null,
): Promise<"sent" | "failed"> {
  const order = ["action_needed", "error", "done", "no_action"];
  items.sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status));
  const { subject, html } = buildReport(mailbox.email_address, items, remaining, run.trigger, appUrl);
  try {
    await client.sendMail(to, subject, html);
    await step(run, "process/report", `Summary email sent to ${to}: "${subject}"`);
    return "sent";
  } catch (err) {
    await logApi("process/report", err, { userId: run.userId, level: "warn", context: { run: run.id, mailbox: mailbox.email_address, to } });
    return "failed";
  }
}

type Db = ReturnType<typeof adminClient>;
interface Mailbox {
  id: string; provider: Provider; auth_type: "oauth" | "app"; connection_id: string | null; email_address: string;
  delegate_address: string | null; credentials_secret_id: string | null; smart_ap_address: string | null;
  mailbox_type: MailboxType; forward_categories: string[]; send_summary: boolean;
  summary_address: string | null;
}
interface ExistingRow {
  id: string; message_id: string; status: string; action_taken: string | null; processed_at: string | null; updated_at: string;
}

// Claims the message for this run: inserts a new row, or (re-check) flips an existing row back to processing
// only if it is still in the status we read, so two runs never work on the same email.
async function claim(db: Db, userId: string, mailbox: Mailbox, msg: MailMessage, existing?: ExistingRow): Promise<string | null> {
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

// What Claude gets from a message's attachments: all names (including files inside attached emails), the PDFs /
// images to read (within size limits), and text: attached emails (.eml / Outlook items) and text files.
interface Gathered { names: string[]; files: AttachmentFile[]; texts: AttachmentText[]; skipped: string[]; total: number }

async function gatherAttachments(client: MailClient, msg: MailMessage) {
  const g: Gathered = { names: [], files: [], texts: [], skipped: [], total: 0 };
  if (!msg.hasAttachments) return g;
  for (const a of await client.listAttachments(msg.id)) {
    g.names.push(a.name);
    const load = () => client.attachmentBytes(msg.id, a.id);
    if (a.isEmail) await addEmail(g, a.name, a.size, load, 1);
    else if (a.isFile) await addFile(g, a.name, a.contentType, a.size, !!a.isInline, load);
  }
  return g;
}

async function addFile(
  g: Gathered, name: string, contentType: string | undefined, size: number, inline: boolean, load: () => Promise<Uint8Array>,
) {
  const type = (contentType ?? "").toLowerCase().split(";")[0].trim();
  const isPdf = type === "application/pdf" || name.toLowerCase().endsWith(".pdf");
  const isImage = IMAGE_TYPES.includes(type) && !inline && size >= MIN_IMAGE_BYTES;
  const isText = TEXT_TYPES.includes(type) || TEXT_EXTENSION.test(name);

  if (isText && !isPdf) {
    if (g.texts.length >= MAX_TEXTS || size > MAX_FILE_BYTES) return g.skipped.push(name);
    const text = new TextDecoder().decode(await load());
    g.texts.push({ name, text: clip(/html?$/.test(type) || /\.html?$/i.test(name) ? htmlToText(text) : text) });
    return;
  }
  if (!isPdf && !isImage) return;
  if (g.files.length >= MAX_FILES || size > MAX_FILE_BYTES || g.total + size > MAX_TOTAL_BYTES) return g.skipped.push(name);
  const bytes = await load();
  g.total += bytes.length;
  g.files.push({ name, mediaType: isPdf ? "application/pdf" : type as AttachmentFile["mediaType"], base64: encodeBase64(bytes) });
}

// An attached email: its headers and text go in as text, its own attachments are handled like the message's.
async function addEmail(g: Gathered, name: string, size: number, load: () => Promise<Uint8Array>, depth: number) {
  if (depth > MAX_EMAIL_DEPTH || g.texts.length >= MAX_TEXTS || size > MAX_EMAIL_BYTES) return g.skipped.push(name);
  const email = await PostalMime.parse(await load());
  const address = (a?: { name: string; address?: string }) => a ? [a.name, a.address && `<${a.address}>`].filter(Boolean).join(" ") : "";
  g.texts.push({
    name,
    text: clip([
      `From: ${address(email.from)}`,
      `To: ${(email.to ?? []).map(address).join(", ")}`,
      `Date: ${email.date ?? ""}`,
      `Subject: ${email.subject ?? ""}`,
      "",
      email.text?.trim() || htmlToText(email.html ?? ""),
    ].join("\n")),
  });

  for (const att of email.attachments) {
    const attName = `${name} › ${att.filename || "attachment"}`;
    const bytes = typeof att.content === "string" ? new TextEncoder().encode(att.content) : new Uint8Array(att.content);
    g.names.push(attName);
    if (isEmailAttachment(att.filename ?? "", att.mimeType)) await addEmail(g, attName, bytes.length, async () => bytes, depth + 1);
    else await addFile(g, attName, att.mimeType, bytes.length, att.disposition === "inline" || !!att.related, async () => bytes);
  }
}

const clip = (text: string) =>
  text.length > MAX_TEXT_CHARS ? `${text.slice(0, MAX_TEXT_CHARS)}\n(Truncated to the first ${MAX_TEXT_CHARS} characters.)` : text;

// Returns the outcome for the summary email, or null if another run already claimed the message.
async function processMessage(
  db: Db, run: Run, mailbox: Mailbox, client: MailClient, msg: MailMessage, existing?: ExistingRow,
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
    const attachments = await gatherAttachments(client, msg);
    await step(run, "process/email",
      `Reading "${subject}" from ${from.address ?? "unknown"}` +
        (attachments.names.length
          ? `: ${attachments.names.length} attachment${attachments.names.length === 1 ? "" : "s"}, ${attachments.files.length + attachments.texts.length} sent to Claude`
          : ": no attachments"),
      { ...email, attachments: attachments.names, sent_to_claude: [...attachments.files, ...attachments.texts].map((f) => f.name), too_large: attachments.skipped });

    const c = await classifyEmail({
      from: from.name ? `${from.name} <${from.address}>` : from.address ?? "unknown",
      subject: msg.subject ?? "",
      received: msg.receivedDateTime,
      attachmentNames: attachments.names,
      files: attachments.files,
      texts: attachments.texts,
      skippedFiles: attachments.skipped,
      body: msg.body?.content ?? "",
    }, mailbox.mailbox_type);
    await step(run, "process/classify", `Classified "${subject}" as ${c.category} (${c.confidence} confidence): ${c.summary}`,
      { ...email, category: c.category, confidence: c.confidence, vendor: c.vendor, amount: c.amount, next_step: c.next_step });

    const { status, action } = await takeAction(run, mailbox, client, msg, c);
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
      ...processedBy(run),
      updated_at: new Date().toISOString(),
    }).eq("id", rowId);
    if (error) throw new Error(error.message);
    await logProcessing(db, run, mailbox, rowId, subject, { status, category: c.category, action });
    return { ...item, category: c.category, summary: c.summary, next_step: c.next_step, action, status };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await logApi("inbox/process", err, { userId, context: { run: run.id, mailbox: mailbox.email_address, ...email } });
    await db.from("processed_emails").update({ status: "error", error: message, ...processedBy(run), updated_at: new Date().toISOString() }).eq("id", rowId);
    await logProcessing(db, run, mailbox, rowId, subject, { status: "error", error: message });
    return { ...item, status: "error", action: message };
  }
}

const processedBy = (run: Run) => ({
  processed_by: run.userId, processed_by_email: run.userEmail, processed_by_name: run.userName, processed_at: new Date().toISOString(),
});

// One email_processing_log row per processed email; a re-check adds another, so earlier outcomes are kept.
// Never throws: a failure to log must not fail the email.
async function logProcessing(
  db: Db, run: Run, mailbox: Mailbox, rowId: string, subject: string,
  outcome: { status: string; category?: string; action?: string | null; error?: string },
) {
  const { error } = await db.from("email_processing_log").insert({
    processed_email_id: rowId,
    mailbox_id: mailbox.id,
    mailbox_address: mailbox.email_address,
    subject,
    owner_user_id: run.userId,
    processed_by: run.userId,
    processed_by_email: run.userEmail,
    processed_by_name: run.userName,
    trigger: run.trigger,
    run_id: run.id,
    status: outcome.status,
    category: outcome.category ?? null,
    action: outcome.action ?? null,
    error: outcome.error ?? null,
  });
  if (error) await logApi("process/log", error.message, { userId: run.userId, level: "warn", context: { run: run.id, email_id: rowId } });
}

// The only automatic action: forward confident emails in the mailbox's forward_categories to its destination address.
// The recipient always comes from settings, never from the email or the model.
async function takeAction(run: Run, mailbox: Mailbox, client: MailClient, msg: MailMessage, c: Classification) {
  const userId = run.userId;
  const label = c.category.replace("_", " ");
  if (c.category === "other") return { status: "no_action", action: null };
  if (!(mailbox.forward_categories ?? []).includes(c.category)) return { status: "action_needed", action: null };

  if (c.confidence === "low") return { status: "action_needed", action: `Not forwarded: not sure this is a ${label} email. Check it, then forward it.` };
  if (!mailbox.smart_ap_address) return { status: "action_needed", action: "Not forwarded: no destination address set for this mailbox." };

  try {
    const details = [c.vendor, c.invoice_number && `#${c.invoice_number}`, c.amount, c.due_date && `due ${c.due_date}`].filter(Boolean).join(", ");
    await client.forwardMessage(msg.id, mailbox.smart_ap_address, `Forwarded automatically by ${APP_NAME}: ${label}${details ? ` (${details})` : ""}.`);
    return { status: "done", action: `Forwarded to ${mailbox.smart_ap_address}` };
  } catch (err) {
    await logApi("inbox/forward", err, { userId, level: "warn", context: { run: run.id, mailbox: mailbox.email_address, subject: msg.subject } });
    const hint = !String(err).includes("403") ? ""
      : mailbox.auth_type === "app" ? " Give the organization's Azure app the Mail.Send application permission." : " Reconnect the mailbox to allow sending.";
    return { status: "action_needed", action: `Forward failed; forward it manually.${hint}` };
  }
}
