// Summary email sent to the mailbox owner after a processing run, rendered with the shared email template.
// Subjects and summaries come from email content, so everything inserted here is HTML-escaped.
import { escapeHtml as esc, renderEmail } from "./email-template.ts";

export interface ReportItem {
  subject: string;
  from: string;
  category: string | null;
  summary: string | null;
  next_step: string | null;
  action: string | null;
  status: string;
  web_link: string | null;
}

const CATEGORY_LABELS: Record<string, string> = {
  invoice: "Invoice",
  utility_bill: "Utility bill",
  vendor_info: "Vendor info",
  statement: "Statement",
  other: "Other",
};

// Section accent colours: action needed / failed / done / no action.
const ACCENT = { action: "#3b5bdb", error: "#a12626", done: "#1d6b3f", none: "#8a94a3" };

function subjectCell(i: ReportItem) {
  const subject = esc(i.subject || "(no subject)");
  const linked = i.web_link?.startsWith("https://")
    ? `<a href="${esc(i.web_link)}" style="color:#1c2230;text-decoration:underline;">${subject}</a>`
    : subject;
  return `<strong>${linked}</strong><br><span style="color:#5f6b7a;font-size:13px;">${esc(i.from)}</span>`;
}

const typeBadge = (i: ReportItem) =>
  `<span style="display:inline-block;padding:2px 8px;border-radius:999px;border:1px solid #dde1e7;font-size:12px;white-space:nowrap;">${
    esc(CATEGORY_LABELS[i.category ?? ""] ?? i.category ?? "")
  }</span>`;

function section(title: string, accent: string, items: ReportItem[], columns: [string, (i: ReportItem) => string][]) {
  if (!items.length) return "";
  const th = columns.map(([h]) =>
    `<th align="left" style="padding:8px;border-bottom:2px solid #dde1e7;font-size:11px;letter-spacing:0.4px;color:#5f6b7a;text-transform:uppercase;">${h}</th>`
  ).join("");
  const rows = items.map((i) =>
    `<tr>${
      columns.map(([, f]) => `<td valign="top" style="padding:10px 8px;border-bottom:1px solid #eef0f3;font-size:14px;line-height:1.45;">${f(i)}</td>`).join("")
    }</tr>`
  ).join("");
  return `
    <h2 style="margin:22px 0 6px;font-size:16px;color:#1c2230;border-left:4px solid ${accent};padding-left:10px;">${title} (${items.length})</h2>
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;"><tr>${th}</tr>${rows}</table>`;
}

function stat(label: string, value: number, color: string) {
  return `<td align="center" style="padding:10px 6px;background:#f5f6f8;border-radius:8px;">
    <div style="font-size:22px;font-weight:700;color:${color};">${value}</div>
    <div style="font-size:12px;color:#5f6b7a;">${label}</div></td>`;
}

export function buildReport(mailbox: string, items: ReportItem[], remaining: number, trigger: string, appUrl: string | null) {
  const needed = items.filter((i) => i.status === "action_needed");
  const done = items.filter((i) => i.status === "done");
  const noAction = items.filter((i) => i.status === "no_action");
  const errors = items.filter((i) => i.status === "error");
  const date = new Date().toISOString().slice(0, 10);
  const s = (n: number) => (n === 1 ? "" : "s");

  const subject = `AP inbox summary – ${date} – ${needed.length} action${s(needed.length)} needed`;

  const stats = `
    <table role="presentation" width="100%" cellspacing="6" cellpadding="0" style="margin:0 0 8px;"><tr>
      ${stat("Action needed", needed.length, ACCENT.action)}
      ${stat("Done", done.length, ACCENT.done)}
      ${stat("No action", noAction.length, ACCENT.none)}
      ${errors.length ? stat("Failed", errors.length, ACCENT.error) : ""}
    </tr></table>`;

  const content = [
    stats,
    remaining
      ? `<p style="margin:10px 0 0;font-size:14px;color:#5f6b7a;">${remaining} more email${s(remaining)} ${remaining === 1 ? "is" : "are"} waiting and will be processed in the next run.</p>`
      : "",
    section("Action needed", ACCENT.action, needed, [
      ["Email", subjectCell],
      ["Type", typeBadge],
      ["Summary", (i) => esc(i.summary)],
      ["Next step", (i) => `<strong>${esc(i.next_step)}</strong>${i.action ? `<br><span style="color:#5f6b7a;font-size:13px;">${esc(i.action)}</span>` : ""}`],
    ]),
    section("Could not process", ACCENT.error, errors, [
      ["Email", subjectCell],
      ["Error", (i) => esc(i.action)],
    ]),
    section("Done automatically", ACCENT.done, done, [
      ["Email", subjectCell],
      ["Type", typeBadge],
      ["What happened", (i) => esc(i.action)],
    ]),
    section("No action needed", ACCENT.none, noAction, [
      ["Email", subjectCell],
      ["Summary", (i) => esc(i.summary)],
    ]),
  ].join("");

  const html = renderEmail({
    title: subject,
    preheader: `${needed.length} action${s(needed.length)} needed, ${done.length} done, ${noAction.length} no action in ${mailbox}.`,
    heading: needed.length ? `${needed.length} email${s(needed.length)} need${needed.length === 1 ? "s" : ""} your attention` : "Your AP inbox is up to date",
    intro: `${trigger === "schedule" ? "Scheduled" : "Manual"} run for <strong>${esc(mailbox)}</strong> on ${esc(date)}: ` +
      `${items.length} email${s(items.length)} processed.`,
    button: appUrl ? { url: `${appUrl}/inbox`, label: "Open the Inbox" } : null,
    content,
    footer: `Sent by SummarizeMail to ${esc(mailbox)}. Classifications are made by AI; check anything unexpected before acting on it.`,
  });
  return { subject, html };
}
