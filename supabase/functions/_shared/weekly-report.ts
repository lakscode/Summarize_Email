// Weekly "forwarded invoices" report of a mailbox (see migrations/*_weekly_invoices_report.sql), separate from the daily
// one (forward-report.ts, unchanged): the invoices and utility bills forwarded in the 7 days before it is sent, by the
// app or manually in Outlook, emailed from the mailbox to its weekly report recipients, with a total per vendor. Sent by
// the scheduler on the day and time set on the Settings page; in a week without any, a short email says so. Every run is kept
// in the report history (reports, kind 'weekly_invoices', shown on the Reports page). Subjects and vendors come from
// email content, so everything inserted is HTML-escaped.
import { APP_NAME } from "./brand.ts";
import { escapeHtml as esc, renderEmail } from "./email-template.ts";
import { logApi, logStep } from "./log.ts";
import { mailClient, type MailboxRow } from "./mail.ts";
import { localDate, nextRun } from "./schedule.ts";
import { adminClient } from "./supabase.ts";

// What the report lists (as the daily report), and the period it covers.
const REPORT_CATEGORIES = ["invoice", "utility_bill"];
const WEEK_MS = 7 * 86_400_000;
const MAX_ROWS = 1000;
// action_taken of a forwarded email: "Forwarded to <address>" (the app) or "Manually forwarded" (process.ts).
const FORWARDED = /^(Forwarded to |Manually forwarded)/;
const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

type Db = ReturnType<typeof adminClient>;

export interface WeeklyReportMailbox extends MailboxRow {
  owner_user_id: string;
  weekly_report_day: number;
  weekly_report_time: string;
  weekly_report_timezone: string;
  weekly_report_addresses: string | null;
}

export const WEEKLY_REPORT_MAILBOX_COLUMNS =
  "id, owner_user_id, provider, auth_type, connection_id, email_address, delegate_address, credentials_secret_id, " +
  "weekly_report_day, weekly_report_time, weekly_report_timezone, weekly_report_addresses, weekly_report_next_at";

// The recipients stored comma-separated in weekly_report_addresses.
export const weeklyRecipients = (value: string | null) => (value ?? "").split(/[\s,;]+/).filter(Boolean);

export const dayName = (day: number) => DAY_NAMES[day] ?? "";

// The next weekly report time after `after`: on `day` (0 = Sunday) at the chosen local time.
export const nextWeeklyReportAt = (day: number, time: string, timezone: string, after: Date) =>
  nextRun({ frequency: "weekdays", days: [day], time: time.slice(0, 5), timezone, anchor: localDate(after, timezone) }, after);

interface Row {
  subject: string | null; from_name: string | null; from_address: string | null; category: string | null;
  vendor: string | null; invoice_number: string | null; amount: string | null; due_date: string | null;
  action_taken: string | null; processed_at: string; web_link: string | null;
}

// Sends one mailbox's weekly report for the 7 days before `now` and, for the scheduled one, schedules the next.
// Returns what happened, also stored in weekly_report_last_result and as a row of the report history; never throws.
// `by`: who pressed "Send now" (null for the scheduled report).
export async function sendWeeklyReport(
  db: Db, m: WeeklyReportMailbox, by: { id: string; name: string | null } | null = null, now = new Date(),
) {
  const since = new Date(now.getTime() - WEEK_MS);
  const to = weeklyRecipients(m.weekly_report_addresses);
  let result: Record<string, unknown>;
  let rows: Row[] = [];
  let subject: string | null = null;
  try {
    const { data, error } = await db.from("processed_emails")
      .select("subject, from_name, from_address, category, vendor, invoice_number, amount, due_date, action_taken, processed_at, web_link")
      .eq("mailbox_id", m.id).eq("status", "done").in("category", REPORT_CATEGORIES)
      .gte("processed_at", since.toISOString()).lt("processed_at", now.toISOString())
      .order("processed_at").limit(MAX_ROWS);
    if (error) throw new Error(error.message);
    rows = (data as Row[]).filter((r) => FORWARDED.test(r.action_taken ?? ""));
    if (!to.length) {
      result = { at: now.toISOString(), error: "No recipients set" };
    } else {
      // Sent also in a week without any (an email saying so), so recipients know the report ran.
      const report = buildWeeklyReport(m, rows, since, now);
      subject = report.subject;
      const sent = await (await mailClient(db, m)).sendMail(to, report.subject, report.html);
      result = { at: now.toISOString(), sent: true, count: rows.length, to, from: sent.from, ...(sent.fallback ? { warning: sent.fallback } : {}) };
      await logStep("weekly-report/sent",
        `Weekly invoices report for ${m.email_address} sent from ${sent.from} to ${to.join(", ")}: ${rows.length || "no invoices"} forwarded`,
        m.owner_user_id, { mailbox: m.email_address, count: rows.length, from: sent.from });
      if (sent.fallback) await logApi("weekly-report", sent.fallback, { userId: m.owner_user_id, level: "warn", context: { mailbox: m.email_address } });
    }
  } catch (err) {
    await logApi("weekly-report", err, { userId: m.owner_user_id, context: { mailbox: m.email_address } });
    result = { at: now.toISOString(), error: err instanceof Error ? err.message : String(err) };
  }

  const { error: historyError } = await db.from("reports").insert({
    kind: "weekly_invoices",
    mailbox_id: m.id,
    mailbox_address: m.email_address,
    owner_user_id: m.owner_user_id,
    trigger: by ? "manual" : "schedule",
    sent_by: by?.id ?? null,
    sent_by_name: by?.name ?? null,
    period_start: since.toISOString(),
    period_end: now.toISOString(),
    // empty: the "no invoices" email was sent.
    status: result.error ? "failed" : rows.length ? "sent" : "empty",
    recipients: to,
    sent_from: (result.from as string | undefined) ?? null,
    subject,
    item_count: rows.length,
    items: rows.map((r) => ({
      subject: r.subject || "(no subject)",
      from: r.from_name ? `${r.from_name} <${r.from_address ?? ""}>` : r.from_address,
      category: r.category, vendor: r.vendor, invoice_number: r.invoice_number, amount: r.amount, due_date: r.due_date,
      forwarded_to: forwardedTo(r), forwarded_at: r.processed_at, web_link: r.web_link,
    })),
    error: (result.error as string | undefined) ?? (result.warning as string | undefined) ?? null,
  });
  if (historyError) await logApi("weekly-report/history", historyError.message, { userId: m.owner_user_id, level: "warn", context: { mailbox: m.email_address } });

  // "Send now" doesn't move the weekly day and time; the scheduled one sets the next week's. A failed scheduled report
  // isn't retried every tick: it waits for the next week.
  await db.from("mailboxes").update({
    ...(by ? {} : { weekly_report_next_at: nextWeeklyReportAt(m.weekly_report_day, m.weekly_report_time, m.weekly_report_timezone, now).toISOString() }),
    ...(result.error ? {} : { weekly_report_last_at: now.toISOString() }),
    weekly_report_last_result: result,
  }).eq("id", m.id);
  return { mailbox: m.email_address, ...result };
}

const LABELS: Record<string, string> = { invoice: "Invoice", utility_bill: "Utility bill" };

// Where an email went: the destination address, or "Manually forwarded".
const forwardedTo = (r: Row) => {
  const action = (r.action_taken ?? "").split(" · ")[0];
  return action.startsWith("Manually forwarded") ? "Manually forwarded" : action.replace(/^Forwarded to /, "");
};

// "1,234.50" / "$1234.5" -> 1234.5; null when the amount isn't a plain number (e.g. "see attached").
function amountValue(amount: string | null): number | null {
  const cleaned = (amount ?? "").replace(/[^0-9.\-]/g, "");
  if (!cleaned || !/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  return Number(cleaned);
}

const money = (n: number) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function buildWeeklyReport(m: WeeklyReportMailbox, rows: Row[], since: Date, until: Date) {
  const tz = m.weekly_report_timezone;
  const when = (d: Date | string) => new Date(d).toLocaleString("en-US", { timeZone: tz, dateStyle: "medium", timeStyle: "short" });
  const day = (d: Date) => new Date(d).toLocaleDateString("en-US", { timeZone: tz, dateStyle: "medium" });
  const n = rows.length;
  const subject = `Weekly forwarded invoices – ${m.email_address} – ${localDate(since, tz)} to ${localDate(until, tz)} – ${n ? `${n} forwarded` : "no invoices"}`;
  if (!n) {
    return {
      subject,
      html: renderEmail({
        title: subject,
        preheader: `No invoices were forwarded from ${m.email_address} this week.`,
        heading: "No invoices to send this week",
        intro: `From <strong>${esc(m.email_address)}</strong>, ${esc(day(since))} to ${esc(day(until))} (${esc(tz)}).`,
        content: `<p style="font-size:14px;line-height:1.5;">No invoices or utility bills were forwarded this week, by ${esc(APP_NAME)} or manually.</p>`,
        footer: `Sent by ${esc(APP_NAME)} for ${esc(m.email_address)} every ${esc(dayName(m.weekly_report_day))}. Change the day, time or recipients on the Settings page.`,
      }),
    };
  }
  const th = (h: string, align = "left") =>
    `<th align="${align}" style="padding:8px;border-bottom:2px solid #dde1e7;font-size:11px;letter-spacing:0.4px;color:#5f6b7a;text-transform:uppercase;">${h}</th>`;
  const cell = (html: string, align = "left") =>
    `<td valign="top" align="${align}" style="padding:10px 8px;border-bottom:1px solid #eef0f3;font-size:14px;line-height:1.45;">${html}</td>`;

  // Per vendor: how many, and the total of the amounts that are plain numbers (others are counted but not added up).
  const vendors = new Map<string, { count: number; total: number; unpriced: number }>();
  for (const r of rows) {
    const key = r.vendor?.trim() || "Unknown vendor";
    const v = vendors.get(key) ?? { count: 0, total: 0, unpriced: 0 };
    v.count++;
    const value = amountValue(r.amount);
    if (value === null) v.unpriced++;
    else v.total += value;
    vendors.set(key, v);
  }
  const vendorRows = [...vendors.entries()].sort((a, b) => b[1].total - a[1].total).map(([name, v]) =>
    `<tr>${cell(esc(name))}${cell(String(v.count), "right")}${cell(esc(money(v.total)) + (v.unpriced ? ` <span style="color:#5f6b7a;font-size:12px;">(+${v.unpriced} without amount)</span>` : ""), "right")}</tr>`
  ).join("");
  const grand = [...vendors.values()].reduce((s, v) => s + v.total, 0);

  const detailRows = rows.map((r) => {
    const title = esc(r.subject || "(no subject)");
    const linked = r.web_link?.startsWith("https://") ? `<a href="${esc(r.web_link)}" style="color:#1c2230;">${title}</a>` : title;
    const from = r.from_name ? `${r.from_name} <${r.from_address ?? ""}>` : r.from_address ?? "";
    return `<tr>${[
      cell(esc(when(r.processed_at))),
      cell(`<strong>${linked}</strong><br><span style="color:#5f6b7a;font-size:13px;">${esc(from)} · ${esc(LABELS[r.category ?? ""] ?? r.category)}</span>`),
      cell(esc(r.vendor)), cell(esc(r.invoice_number)), cell(esc(r.amount)), cell(esc(r.due_date)), cell(esc(forwardedTo(r))),
    ].join("")}</tr>`;
  }).join("");

  const table = (head: string, body: string) =>
    `<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>${head}</tr>${body}</table>`;
  const html = renderEmail({
    title: subject,
    preheader: `${n} invoice${n === 1 ? "" : "s"} forwarded from ${m.email_address} this week.`,
    heading: `${n} invoice${n === 1 ? "" : "s"} forwarded this week`,
    intro: `From <strong>${esc(m.email_address)}</strong>, ${esc(day(since))} to ${esc(day(until))} (${esc(tz)}).`,
    content:
      `<h2 style="margin:22px 0 6px;font-size:16px;color:#1c2230;">By vendor</h2>` +
      table(th("Vendor") + th("Invoices", "right") + th("Total", "right"),
        vendorRows + `<tr>${cell("<strong>All vendors</strong>")}${cell(`<strong>${n}</strong>`, "right")}${cell(`<strong>${esc(money(grand))}</strong>`, "right")}</tr>`) +
      `<h2 style="margin:22px 0 6px;font-size:16px;color:#1c2230;">Invoices</h2>` +
      table(["Forwarded", "Email", "Vendor", "Invoice #", "Amount", "Due", "Forwarded to"].map((h) => th(h)).join(""), detailRows) +
      (n >= MAX_ROWS ? `<p style="font-size:13px;color:#5f6b7a;">Only the first ${MAX_ROWS} are listed.</p>` : ""),
    footer: `Sent by ${esc(APP_NAME)} for ${esc(m.email_address)} every ${esc(dayName(m.weekly_report_day))}. ` +
      `Totals add up the amounts as written on the invoices, whatever their currency. Change the day, time or recipients on the Settings page.`,
  });
  return { subject, html };
}
