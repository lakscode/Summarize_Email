// Build the one summary email the mailbox owner receives after each run.
import type { SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import type { Category, ItemStatus, Mailbox, ReconStatus } from "./types.ts";
import { escapeHtml as h } from "./util.ts";
import { money } from "./process.ts";

interface EmailRow {
  id: string;
  run_id: string | null;
  category: Category;
  status: ItemStatus;
  vendor_name: string | null;
  from_address: string | null;
  subject: string | null;
  received_at: string | null;
  invoice_number: string | null;
  amount: number | null;
  currency: string | null;
  due_date: string | null;
  summary: string | null;
  action_taken: string | null;
  employee_action: string | null;
  flags: string[];
  error: string | null;
  statement_lines: {
    invoice_number: string | null;
    invoice_date: string | null;
    amount: number | null;
    recon_status: ReconStatus;
    note: string | null;
  }[];
}

const LABEL: Record<Category, string> = {
  invoice: "Invoice",
  utility_bill: "Portal bill",
  vendor_info: "Vendor info",
  statement: "Account statement",
  other: "Other",
};

const fmtDate = (s: string | null) =>
  s ? new Date(s).toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" }) : "";

const byDue = (a: EmailRow, b: EmailRow) => (a.due_date ?? "9999").localeCompare(b.due_date ?? "9999");

export async function buildReport(db: SupabaseClient, mailbox: Mailbox, runId: string) {
  const cols =
    "id, run_id, category, status, vendor_name, from_address, subject, received_at, invoice_number, amount, currency, due_date, summary, action_taken, employee_action, flags, error, statement_lines(invoice_number, invoice_date, amount, recon_status, note)";
  const [{ data: thisRun, error: e1 }, { data: carried, error: e2 }] = await Promise.all([
    db.from("emails").select(cols).eq("run_id", runId).neq("status", "skipped"),
    db.from("emails").select(cols).eq("mailbox_id", mailbox.id).neq("run_id", runId)
      .in("status", ["action_needed", "flagged"]).is("completed_at", null),
  ]);
  if (e1 || e2) throw new Error(`Loading report data failed: ${(e1 ?? e2)!.message}`);
  const rows = (thisRun ?? []) as EmailRow[];
  const open = (carried ?? []) as EmailRow[];

  const counts: Record<Category, number> = { invoice: 0, utility_bill: 0, vendor_info: 0, statement: 0, other: 0 };
  for (const r of rows) counts[r.category]++;

  const all = [...rows, ...open];
  const isOpen = (r: EmailRow) => r.status === "action_needed" || r.status === "flagged";
  const sections: { title: string; items: EmailRow[] }[] = [
    { title: "⚠️ Needs attention", items: all.filter((r) => r.status === "flagged" || (r.category === "other" && isOpen(r))).sort(byDue) },
    { title: "📑 Statements - missing invoices", items: all.filter((r) => r.category === "statement" && isOpen(r) && r.status !== "flagged").sort(byDue) },
    { title: "🔌 Portal bills to download", items: all.filter((r) => r.category === "utility_bill" && isOpen(r) && r.status !== "flagged").sort(byDue) },
    { title: "📝 Review and file", items: all.filter((r) => r.category === "vendor_info" && isOpen(r) && r.status !== "flagged").sort(byDue) },
    { title: "🧾 Invoices", items: all.filter((r) => r.category === "invoice" && r.status !== "flagged" && (isOpen(r) || r.run_id === runId)).sort(byDue) },
  ];
  const actionCount = all.filter(isOpen).length;

  let n = 0;
  const itemHtml = (r: EmailRow) => {
    n++;
    const carriedOver = r.run_id !== runId ? ` <span style="color:#a15c00">(still open from ${fmtDate(r.received_at)})</span>` : "";
    const missing = r.statement_lines?.filter((l) => l.recon_status === "missing" || l.recon_status === "amount_mismatch") ?? [];
    const table = missing.length
      ? `<table cellpadding="4" style="border-collapse:collapse;margin:6px 0;font-size:13px">
          <tr style="background:#f3f4f6"><th align="left">Invoice #</th><th align="left">Date</th><th align="right">Amount</th><th align="left">Status</th></tr>
          ${missing.map((l) => `<tr><td>${h(l.invoice_number)}</td><td>${h(l.invoice_date)}</td><td align="right">${money(l.amount)}</td>
            <td>${l.recon_status === "missing" ? "Missing" : `Amount mismatch - ${h(l.note)}`}</td></tr>`).join("")}
        </table>`
      : "";
    return `<li style="margin-bottom:12px">
      <b>${h(LABEL[r.category])} - ${h(r.vendor_name ?? r.from_address)}</b>${carriedOver}<br>
      <span style="color:#555">${h(r.subject)} · received ${fmtDate(r.received_at)}${r.amount != null ? ` · ${money(r.amount, r.currency)}` : ""}${r.due_date ? ` · <b>due ${fmtDate(r.due_date)}</b>` : ""}</span><br>
      ${r.action_taken ? `Done: ${h(r.action_taken)}<br>` : ""}
      ${r.employee_action ? `<b>Your next step:</b> ${h(r.employee_action)}` : ""}
      ${r.flags?.length ? `<br><span style="color:#b42318">Signals: ${h(r.flags.join("; "))}</span>` : ""}
      ${table}
    </li>`;
  };

  const body = sections
    .filter((s) => s.items.length)
    .map((s) => `<h3 style="margin:20px 0 8px">${s.title}</h3><ol start="${n + 1}" style="padding-left:20px">${s.items.map(itemHtml).join("")}</ol>`)
    .join("");

  const noAction = rows.filter((r) => r.status === "info");
  const errors = rows.filter((r) => r.error);
  const date = new Date().toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });

  const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111;max-width:760px">
    <h2 style="margin:0 0 4px">AP Inbox Action Report - ${h(date)}</h2>
    <p style="margin:0 0 12px;color:#555">${h(mailbox.email_address)} · processed ${rows.length} new email(s):
      ${counts.invoice} invoices, ${counts.utility_bill} portal bills, ${counts.vendor_info} vendor info, ${counts.statement} statements, ${counts.other} other.</p>
    ${actionCount ? `<p><b>${actionCount} item(s) need you.</b> Soonest due dates are listed first in each section.</p>` : `<p><b>Nothing needs your attention.</b></p>`}
    ${body}
    ${noAction.length ? `<h3 style="margin:20px 0 8px">Other (no action)</h3><ul>${noAction.map((r) => `<li>${h(r.from_address)} - ${h(r.subject)}</li>`).join("")}</ul>` : ""}
    ${errors.length ? `<h3 style="margin:20px 0 8px">Could not process automatically</h3><ul>${errors.map((r) => `<li>${h(r.from_address)} - ${h(r.subject)}: ${h(r.error)}</li>`).join("")}</ul>` : ""}
  </div>`;

  return {
    subject: `AP Inbox Action Report - ${date} (${actionCount} action item${actionCount === 1 ? "" : "s"})`,
    html,
    counts,
    actionCount,
  };
}
