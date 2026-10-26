// Decide and carry out the action for one email, then store it as an action item.
import type { SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import { classifyEmail } from "./classify.ts";
import { reconcileStatement, type ReconResult } from "./reconcile.ts";
import { isFlagged, safetyCheck, trustedPortalUrl } from "./safety.ts";
import type { Category, Classification, ItemStatus, Mailbox, MailMessage, MailProvider } from "./types.ts";
import { normalizeInvoiceNumber } from "./util.ts";

export function money(amount: number | null | undefined, currency?: string | null): string {
  if (amount == null) return "amount not shown";
  const cur = currency && /^[A-Z]{3}$/.test(currency) ? currency : "USD";
  try {
    return new Intl.NumberFormat("en-US", { style: "currency", currency: cur }).format(amount);
  } catch {
    return amount.toFixed(2);
  }
}

function isoDate(s: string | null | undefined): string | null {
  return s && /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
}

interface Decision {
  status: ItemStatus;
  actionTaken: string | null;
  employeeAction: string | null;
  label: boolean; // tag the email as processed in the mailbox
  forwardedAt: string | null;
}

export async function processEmail(
  db: SupabaseClient,
  provider: MailProvider,
  mailbox: Mailbox,
  runId: string,
  msg: MailMessage,
): Promise<void> {
  const base = {
    mailbox_id: mailbox.id,
    run_id: runId,
    provider_message_id: msg.id,
    thread_id: msg.threadId,
    from_address: msg.from.address,
    from_name: msg.from.name,
    subject: msg.subject,
    received_at: msg.receivedAt.toISOString(),
    attachment_names: msg.attachments.map((a) => a.name),
  };

  const safety = safetyCheck(msg);
  let c: Classification;
  let unreadable: string[];
  try {
    ({ classification: c, unreadable } = await classifyEmail(provider, msg, {
      portalVendors: mailbox.portal_vendors,
      managedEntities: mailbox.managed_entities,
      safety,
    }));
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await upsertEmail(db, {
      ...base,
      category: "other",
      status: "action_needed",
      summary: "Could not be classified automatically.",
      employee_action: "Open this email and handle it manually - automatic classification failed.",
      error: message,
      flags: [...safety.strong, ...safety.weak],
    });
    return;
  }

  const vendor = c.vendor_name ?? msg.from.name ?? msg.from.address;
  const flagged = isFlagged(safety, c.fraud_risk) || c.other_subtype === "bank_detail_change";
  const flags = [...new Set([...safety.strong, ...(flagged ? safety.weak : []), ...c.fraud_signals])];
  const fraudLine =
    `⚠️ Possible fraud - verify with ${vendor} using a phone number you already have, not one from this email.`;
  const extracted: Record<string, unknown> = { ...c, unreadable_attachments: unreadable };

  let recon: ReconResult | null = null;
  if (c.statement?.lines.length) {
    recon = await reconcileStatement(db, mailbox, vendor, c.statement.lines);
  }

  const d = await decide(db, provider, mailbox, msg, c, { vendor, flagged, fraudLine, recon, extracted });

  if (unreadable.length && d.status !== "flagged") {
    d.employeeAction = [d.employeeAction, `Check attachment(s) manually: ${unreadable.join(", ")}.`]
      .filter(Boolean).join(" ");
    if (d.status === "info" || d.status === "done") d.status = "action_needed";
  }

  let labeledAt: string | null = null;
  if (d.label) {
    try {
      await provider.addLabel(msg.id, mailbox.processed_label);
      labeledAt = new Date().toISOString();
    } catch (err) {
      console.warn(`Could not label ${msg.id}:`, err);
    }
  }

  const amount = c.invoice?.amount ?? c.utility?.amount ?? c.statement?.closing_balance ?? null;
  const emailId = await upsertEmail(db, {
    ...base,
    category: c.category,
    confidence: c.confidence,
    vendor_name: vendor,
    invoice_number: c.invoice?.invoice_number ?? null,
    amount,
    currency: c.invoice?.currency ?? null,
    due_date: isoDate(c.invoice?.due_date ?? c.utility?.due_date),
    summary: c.one_line_summary,
    reasoning: c.reasoning,
    extracted,
    flags,
    status: d.status,
    action_taken: d.actionTaken,
    employee_action: d.employeeAction,
    forwarded_at: d.forwardedAt,
    labeled_at: labeledAt,
  });

  if (recon) {
    await db.from("statement_lines").delete().eq("email_id", emailId);
    const { error } = await db.from("statement_lines").insert(
      recon.lines.map((l) => ({
        email_id: emailId,
        invoice_number: l.invoice_number,
        invoice_date: isoDate(l.invoice_date),
        amount: l.amount,
        balance: l.balance,
        vendor_status: l.status,
        recon_status: l.recon_status,
        matched_ref: l.matched_ref,
        note: l.note,
      })),
    );
    if (error) throw new Error(`Saving statement lines failed: ${error.message}`);
  }
}

async function decide(
  db: SupabaseClient,
  provider: MailProvider,
  mailbox: Mailbox,
  msg: MailMessage,
  c: Classification,
  x: { vendor: string; flagged: boolean; fraudLine: string; recon: ReconResult | null; extracted: Record<string, unknown> },
): Promise<Decision> {
  const category: Category = c.category;
  const reconLine = x.recon ? statementAction(x.vendor, x.recon) : null;

  if (x.flagged) {
    return {
      status: "flagged",
      actionTaken: category === "invoice" ? "Not forwarded to Smart AP - flagged for verification." : null,
      employeeAction: `${x.fraudLine} Do not pay, forward or update vendor records until verified.`,
      label: false,
      forwardedAt: null,
    };
  }

  switch (category) {
    case "invoice": {
      const inv = c.invoice;
      const num = inv?.invoice_number ?? "(no number)";
      const desc = `${x.vendor} invoice ${num} (${money(inv?.amount, inv?.currency)}${inv?.due_date ? `, due ${inv.due_date}` : ""})`;
      const extra = reconLine ? ` Also: ${reconLine.text}` : "";

      if (mailbox.managed_entities.length && inv?.billed_entity) {
        const billed = inv.billed_entity.toLowerCase();
        const ours = mailbox.managed_entities.some((e) => billed.includes(e.toLowerCase()) || e.toLowerCase().includes(billed));
        if (!ours) {
          return {
            status: "action_needed",
            actionTaken: "Not forwarded - billed entity is not on the managed list.",
            employeeAction: `${desc} is billed to "${inv.billed_entity}", which is not an entity we manage. Confirm before sending it to Smart AP.${extra}`,
            label: false,
            forwardedAt: null,
          };
        }
      }

      if (!mailbox.smart_ap_address || !mailbox.auto_forward_invoices) {
        const how = mailbox.smart_ap_address
          ? `Forward it to Smart AP (${mailbox.smart_ap_address}) - use "Send to Smart AP" in the dashboard.`
          : "Forward it to the Smart AP inbox (set the Smart AP address in settings to enable one-click forwarding).";
        return {
          status: "action_needed",
          actionTaken: null,
          employeeAction: `${desc}: ${how}${extra}`,
          label: false,
          forwardedAt: null,
        };
      }

      const already = await alreadyForwarded(db, mailbox, x.vendor, inv?.invoice_number ?? null);
      if (already) {
        return {
          status: reconLine?.needsAction ? "action_needed" : "done",
          actionTaken: `Already sent to Smart AP on ${already.slice(0, 10)} - not sent again.`,
          employeeAction: reconLine?.needsAction ? reconLine.text : null,
          label: true,
          forwardedAt: null,
        };
      }

      await provider.forward(msg, mailbox.smart_ap_address, `Invoice ${num} from ${x.vendor} for processing.`);
      return {
        status: reconLine?.needsAction ? "action_needed" : "done",
        actionTaken: `Forwarded to Smart AP (${mailbox.smart_ap_address}).`,
        employeeAction: `Confirm ${desc} appears in Smart AP.${extra}`,
        label: true,
        forwardedAt: new Date().toISOString(),
      };
    }

    case "utility_bill": {
      const u = c.utility;
      const portal = trustedPortalUrl(u?.portal_url ?? null, msg.from.address, mailbox.portal_vendors);
      if (u) x.extracted.utility = { ...u, portal_url: portal.url };
      const name = u?.provider ?? x.vendor;
      const acct = u?.account_numbers.length ? ` for account ${u.account_numbers.join(", ")}` : "";
      const bits = [u?.billing_period, u?.amount != null ? money(u.amount) : null, u?.due_date ? `due ${u.due_date}` : null]
        .filter(Boolean).join(", ");
      return {
        status: "action_needed",
        actionTaken: null,
        employeeAction:
          `Log into the ${name} portal${portal.url ? ` (${portal.url})` : ""}, download the invoice${acct}` +
          `${bits ? ` (${bits})` : ""} and send it to Smart AP for processing.` +
          (portal.warning ? ` ⚠️ ${portal.warning}` : ""),
        label: false, // stays unlabeled until the employee marks it done
        forwardedAt: null,
      };
    }

    case "vendor_info":
      return {
        status: "action_needed",
        actionTaken: null,
        employeeAction: `${c.vendor_info?.summary ?? c.one_line_summary} Please review and keep it for later; no invoice to process yet.` +
          (c.vendor_info?.key_dates.length ? ` Key dates: ${c.vendor_info.key_dates.join("; ")}.` : ""),
        label: true,
        forwardedAt: null,
      };

    case "statement": {
      if (!x.recon) {
        return {
          status: "action_needed",
          actionTaken: null,
          employeeAction: `Could not read the invoice lines on the ${x.vendor} statement. Check it manually against the AP system.`,
          label: true,
          forwardedAt: null,
        };
      }
      return {
        status: reconLine!.needsAction ? "action_needed" : "done",
        actionTaken: `Reconciled ${x.recon.lines.length} statement line(s). ${x.recon.sourceNote}`,
        employeeAction: reconLine!.text,
        label: true,
        forwardedAt: null,
      };
    }

    default: {
      const needs = c.urgent || c.confidence === "low" ||
        ["past_due", "dispute_or_credit", "unclear"].includes(c.other_subtype ?? "");
      return {
        status: needs ? "action_needed" : "info",
        actionTaken: null,
        employeeAction: c.other_subtype === "not_ap" ? "No AP action." : c.employee_action,
        label: !needs,
        forwardedAt: null,
      };
    }
  }
}

function statementAction(vendor: string, recon: ReconResult): { text: string; needsAction: boolean } {
  const missing = recon.lines.filter((l) => l.recon_status === "missing");
  const mismatch = recon.lines.filter((l) => l.recon_status === "amount_mismatch");
  const fmt = (l: (typeof recon.lines)[number]) =>
    `${l.invoice_number ?? "(no number)"}${l.invoice_date ? ` ${l.invoice_date}` : ""} ${money(l.amount)}`;
  const parts: string[] = [];
  if (missing.length) {
    parts.push(`Request copies of these missing invoices from ${vendor} and send them to Smart AP: ${missing.map(fmt).join("; ")}.`);
  }
  if (mismatch.length) {
    parts.push(`Check amount differences: ${mismatch.map((l) => `${l.invoice_number} (${l.note})`).join("; ")}.`);
  }
  if (!parts.length) {
    return { text: `All ${recon.lines.length} invoices on the ${vendor} statement are already recorded. Nothing to request.`, needsAction: false };
  }
  return { text: `${parts.join(" ")} (${recon.sourceNote})`, needsAction: true };
}

async function alreadyForwarded(db: SupabaseClient, mailbox: Mailbox, vendor: string, invoiceNumber: string | null): Promise<string | null> {
  const norm = normalizeInvoiceNumber(invoiceNumber);
  if (!norm) return null;
  const { data } = await db
    .from("emails")
    .select("invoice_number, vendor_name, forwarded_at")
    .eq("mailbox_id", mailbox.id)
    .eq("category", "invoice")
    .not("forwarded_at", "is", null);
  const hit = (data ?? []).find(
    (r) => normalizeInvoiceNumber(r.invoice_number) === norm &&
      (r.vendor_name ?? "").toLowerCase().split(" ")[0] === vendor.toLowerCase().split(" ")[0],
  );
  return hit?.forwarded_at ?? null;
}

async function upsertEmail(db: SupabaseClient, row: Record<string, unknown>): Promise<string> {
  const { data, error } = await db
    .from("emails")
    .upsert(row, { onConflict: "mailbox_id,provider_message_id" })
    .select("id")
    .single();
  if (error) throw new Error(`Saving email failed: ${error.message}`);
  return data.id as string;
}
