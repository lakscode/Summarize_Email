// Daily "forwarded invoices" report of a mailbox (see migrations/*_forwarded_invoices_report.sql): the invoices and
// utility bills forwarded since the previous report, by the app or manually in Outlook, emailed from the mailbox to
// its report recipients. Sent by the scheduler at the time set on the Settings page; nothing is sent when nothing was
// forwarded. Subjects and vendors come from email content, so everything inserted is HTML-escaped.
import { APP_NAME } from "./brand.ts";
import { escapeHtml as esc, renderEmail } from "./email-template.ts";
import { logApi, logStep } from "./log.ts";
import { mailClient, type MailboxRow } from "./mail.ts";
import { localDate, nextRun } from "./schedule.ts";
import { adminClient } from "./supabase.ts";

// What the report lists, and how far back a first report (or one after a long pause) looks.
const REPORT_CATEGORIES = ["invoice", "utility_bill"];
const MAX_LOOKBACK_MS = 7 * 86_400_000;
const MAX_ROWS = 500;
// action_taken of a forwarded email: "Forwarded to <address>" (the app) or "Manually forwarded" (process.ts).
const FORWARDED = /^(Forwarded to |Manually forwarded)/;

type Db = ReturnType<typeof adminClient>;

export interface ReportMailbox extends MailboxRow {
  owner_user_id: string;
  forward_report_time: string;
  forward_report_timezone: string;
  forward_report_addresses: string | null;
  forward_report_last_at: string | null;
}

export const REPORT_MAILBOX_COLUMNS =
  "id, owner_user_id, provider, auth_type, connection_id, email_address, delegate_address, credentials_secret_id, " +
  "forward_report_time, forward_report_timezone, forward_report_addresses, forward_report_last_at, forward_report_next_at";

// The recipients stored comma-separated in forward_report_addresses.
export const reportRecipients = (value: string | null) => (value ?? "").split(/[\s,;]+/).filter(Boolean);

// The next report time after `after`: every day at the chosen local time.
export const nextReportAt = (time: string, timezone: string, after: Date) =>
  nextRun({ frequency: "daily", days: [], time: time.slice(0, 5), timezone, anchor: localDate(after, timezone) }, after);

interface Row {
  subject: string | null; from_name: string | null; from_address: string | null; category: string | null;
  vendor: string | null; invoice_number: string | null; amount: string | null; due_date: string | null;
  action_taken: string | null; processed_at: string; web_link: string | null;
}

// Sends one mailbox's report for the time since its previous one and schedules the next. Returns what happened, also
// stored in forward_report_last_result; never throws.
export async function sendForwardReport(db: Db, m: ReportMailbox, now = new Date()) {
  const since = new Date(Math.max(
    m.forward_report_last_at ? new Date(m.forward_report_last_at).getTime() : 0,
    now.getTime() - (m.forward_report_last_at ? MAX_LOOKBACK_MS : 86_400_000),
  ));
  const next = nextReportAt(m.forward_report_time, m.forward_report_timezone, now);
  let result: Record<string, unknown>;
  try {
    const { data, error } = await db.from("processed_emails")
      .select("subject, from_name, from_address, category, vendor, invoice_number, amount, due_date, action_taken, processed_at, web_link")
      .eq("mailbox_id", m.id).eq("status", "done").in("category", REPORT_CATEGORIES)
      .gte("processed_at", since.toISOString()).lt("processed_at", now.toISOString())
      .order("processed_at").limit(MAX_ROWS);
    if (error) throw new Error(error.message);
    const rows = (data as Row[]).filter((r) => FORWARDED.test(r.action_taken ?? ""));
    const to = reportRecipients(m.forward_report_addresses);
    if (!to.length) {
      result = { at: now.toISOString(), error: "No recipients set" };
    } else if (!rows.length) {
      result = { at: now.toISOString(), sent: false, count: 0 };
    } else {
      const { subject, html } = buildForwardReport(m, rows, since, now);
      await (await mailClient(db, m)).sendMail(to, subject, html);
      result = { at: now.toISOString(), sent: true, count: rows.length, to };
      await logStep("forward-report/sent", `Forwarded invoices report for ${m.email_address} sent to ${to.join(", ")}: ${rows.length} forwarded`,
        m.owner_user_id, { mailbox: m.email_address, count: rows.length });
    }
  } catch (err) {
    await logApi("forward-report", err, { userId: m.owner_user_id, context: { mailbox: m.email_address } });
    result = { at: now.toISOString(), error: err instanceof Error ? err.message : String(err) };
  }
  // A failed report isn't retried every tick: the next one covers the same period (since the last successful one).
  await db.from("mailboxes").update({
    forward_report_next_at: next.toISOString(),
    ...(result.error ? {} : { forward_report_last_at: now.toISOString() }),
    forward_report_last_result: result,
  }).eq("id", m.id);
  return { mailbox: m.email_address, ...result };
}

const LABELS: Record<string, string> = { invoice: "Invoice", utility_bill: "Utility bill" };

function buildForwardReport(m: ReportMailbox, rows: Row[], since: Date, until: Date) {
  const tz = m.forward_report_timezone;
  const when = (d: Date | string) => new Date(d).toLocaleString("en-US", { timeZone: tz, dateStyle: "medium", timeStyle: "short" });
  const date = localDate(until, tz);
  const n = rows.length;
  const subject = `Forwarded invoices – ${m.email_address} – ${date} – ${n} forwarded`;
  const head = ["Forwarded", "Email", "Vendor", "Invoice #", "Amount", "Due", "Forwarded to"].map((h) =>
    `<th align="left" style="padding:8px;border-bottom:2px solid #dde1e7;font-size:11px;letter-spacing:0.4px;color:#5f6b7a;text-transform:uppercase;">${h}</th>`
  ).join("");
  const cell = (html: string) => `<td valign="top" style="padding:10px 8px;border-bottom:1px solid #eef0f3;font-size:14px;line-height:1.45;">${html}</td>`;
  const body = rows.map((r) => {
    const title = esc(r.subject || "(no subject)");
    const linked = r.web_link?.startsWith("https://") ? `<a href="${esc(r.web_link)}" style="color:#1c2230;">${title}</a>` : title;
    const from = r.from_name ? `${r.from_name} <${r.from_address ?? ""}>` : r.from_address ?? "";
    const forwardedTo = r.action_taken === "Manually forwarded" ? "Manually forwarded" : (r.action_taken ?? "").replace(/^Forwarded to /, "");
    return `<tr>${[
      cell(esc(when(r.processed_at))),
      cell(`<strong>${linked}</strong><br><span style="color:#5f6b7a;font-size:13px;">${esc(from)} · ${esc(LABELS[r.category ?? ""] ?? r.category)}</span>`),
      cell(esc(r.vendor)), cell(esc(r.invoice_number)), cell(esc(r.amount)), cell(esc(r.due_date)), cell(esc(forwardedTo)),
    ].join("")}</tr>`;
  }).join("");
  const html = renderEmail({
    title: subject,
    preheader: `${n} invoice${n === 1 ? "" : "s"} forwarded from ${m.email_address}.`,
    heading: `${n} invoice${n === 1 ? "" : "s"} forwarded`,
    intro: `From <strong>${esc(m.email_address)}</strong> between ${esc(when(since))} and ${esc(when(until))} (${esc(tz)}).`,
    content: `<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>${head}</tr>${body}</table>` +
      (n >= MAX_ROWS ? `<p style="font-size:13px;color:#5f6b7a;">Only the first ${MAX_ROWS} are listed.</p>` : ""),
    footer: `Sent by ${esc(APP_NAME)} for ${esc(m.email_address)}. Change the time or recipients on the Settings page.`,
  });
  return { subject, html };
}
