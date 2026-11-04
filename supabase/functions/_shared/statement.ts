// Account statements: each invoice number on the statement is looked up among the invoices (and utility bills) the
// mailbox has processed, to tell whether it was already forwarded (and when), received but not forwarded, or never
// received. Used when a statement is processed (process.ts) and when someone checks again in the Inbox (inbox/index.ts).
import { adminClient } from "./supabase.ts";

type Db = ReturnType<typeof adminClient>;

export interface StatementLine {
  invoice_number: string;
  // forwarded: sent on (by the app, or manually in Outlook); received: processed but not forwarded; not_found: no email
  result: "forwarded" | "received" | "not_found";
  forwarded_at?: string | null;
  received_at?: string | null;
  email_id?: string;
  subject?: string | null;
  vendor?: string | null;
  email_status?: string;
}

// Most invoices of one mailbox compared; the most recent.
const MAX_INVOICES = 5000;
const FORWARDED = ["Forwarded to", "Manually forwarded"];
const isForwarded = (action: string | null) => FORWARDED.some((f) => (action ?? "").startsWith(f));

// "INV-00123" and "inv 00123" are the same invoice; so are "INV-00123" and "123" (same digits, at least 4 of them, so
// short numbers don't match by chance).
const plain = (s: string) => s.toUpperCase().replace(/[^A-Z0-9]/g, "");
const digits = (s: string) => s.replace(/\D/g, "").replace(/^0+/, "");
function sameInvoice(a: string, b: string) {
  if (plain(a) && plain(a) === plain(b)) return true;
  const d = digits(a);
  return d.length >= 4 && d === digits(b);
}
// An invoice email's invoice_number can hold several ("1043, 1047").
const numbersOf = (value: string) => value.split(/\s*(?:[,;&]|\band\b)\s*/i).map((s) => s.replace(/^#/, "").trim()).filter(Boolean);

export async function checkStatement(db: Db, mailboxId: string, statementId: string, invoices: string[]): Promise<StatementLine[]> {
  const wanted = [...new Set(invoices.map((s) => s.replace(/^#/, "").trim()).filter(Boolean))];
  if (!wanted.length) return [];
  const { data: rows, error } = await db.from("processed_emails")
    .select("id, invoice_number, vendor, subject, status, action_taken, received_at, processed_at")
    .eq("mailbox_id", mailboxId).in("category", ["invoice", "utility_bill"]).not("invoice_number", "is", null).neq("id", statementId)
    .order("received_at", { ascending: false }).limit(MAX_INVOICES);
  if (error) throw new Error(error.message);

  // For each invoice: a forwarded email first, else the most recent one.
  const matches = new Map<string, typeof rows[number]>();
  for (const n of wanted) {
    const found = rows.filter((r) => numbersOf(r.invoice_number).some((m) => sameInvoice(n, m)));
    const best = found.find((r) => isForwarded(r.action_taken)) ?? found[0];
    if (best) matches.set(n, best);
  }

  // When each was forwarded: the first forwarding in its processing history.
  const forwardedIds = [...matches.values()].filter((r) => isForwarded(r.action_taken)).map((r) => r.id);
  const forwardedAt = new Map<string, string>();
  if (forwardedIds.length) {
    const { data: log, error: logError } = await db.from("email_processing_log").select("processed_email_id, processed_at")
      .in("processed_email_id", forwardedIds).or(FORWARDED.map((f) => `action.like."${f}%"`).join(","))
      .order("processed_at", { ascending: true });
    if (logError) throw new Error(logError.message);
    for (const l of log) if (!forwardedAt.has(l.processed_email_id)) forwardedAt.set(l.processed_email_id, l.processed_at);
  }

  return wanted.map((n): StatementLine => {
    const r = matches.get(n);
    if (!r) return { invoice_number: n, result: "not_found" };
    const forwarded = isForwarded(r.action_taken);
    return {
      invoice_number: n,
      result: forwarded ? "forwarded" : "received",
      // Forwarded before the processing history was kept: when the email was last processed.
      forwarded_at: forwarded ? forwardedAt.get(r.id) ?? r.processed_at : null,
      received_at: r.received_at,
      email_id: r.id,
      subject: r.subject,
      vendor: r.vendor,
      email_status: r.status,
    };
  });
}
