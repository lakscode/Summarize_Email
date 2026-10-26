// Category D: check every statement line against the AP ledger, then against invoices
// already received by email, and report what is missing.
import type { SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import type { Mailbox, ReconStatus, StatementLine } from "./types.ts";
import { normalizeInvoiceNumber } from "./util.ts";

export interface ReconciledLine extends StatementLine {
  recon_status: ReconStatus;
  matched_ref: string | null;
  note: string | null;
}

export interface ReconResult {
  lines: ReconciledLine[];
  sourceNote: string; // tells the employee how far to trust the result
}

interface Match {
  norm: string;
  source: "ledger" | "email";
  ref: string;
  vendor_name: string | null;
  amount: number | null;
}

function vendorKey(name: string | null): string {
  return (name ?? "").toLowerCase().replace(/[^a-z0-9 ]/g, "").split(" ").find((w) => w.length > 2) ?? "";
}

export async function reconcileStatement(
  db: SupabaseClient,
  mailbox: Mailbox,
  vendor: string | null,
  lines: StatementLine[],
): Promise<ReconResult> {
  const numbers = [...new Set(lines.map((l) => normalizeInvoiceNumber(l.invoice_number)).filter(Boolean))];
  const { data, error } = numbers.length
    ? await db.rpc("find_invoices", { p_owner: mailbox.owner_user_id, p_mailbox: mailbox.id, p_numbers: numbers })
    : { data: [], error: null };
  if (error) throw new Error(`Invoice lookup failed: ${error.message}`);
  const matches = (data ?? []) as Match[];

  const { count: ledgerCount } = await db
    .from("ap_ledger")
    .select("id", { count: "exact", head: true })
    .eq("owner_user_id", mailbox.owner_user_id);
  const sourceNote = ledgerCount
    ? "Checked: AP ledger and invoices received in this mailbox."
    : "Checked: this mailbox only - no AP ledger has been imported, so verify the result in the AP system.";

  const vKey = vendorKey(vendor);
  const out: ReconciledLine[] = lines.map((line) => {
    const norm = normalizeInvoiceNumber(line.invoice_number);
    if (!norm) {
      return { ...line, recon_status: "missing", matched_ref: null, note: "No invoice number on the statement line" };
    }
    const candidates = matches.filter((m) => m.norm === norm);
    // Prefer matches for the same vendor; invoice numbers are only unique per vendor.
    const sameVendor = candidates.filter((m) => !vKey || vendorKey(m.vendor_name) === vKey);
    const pool = sameVendor.length ? sameVendor : candidates;
    const ledger = pool.find((m) => m.source === "ledger");
    const email = pool.find((m) => m.source === "email");
    const hit = ledger ?? email;
    if (!hit) return { ...line, recon_status: "missing", matched_ref: null, note: null };

    const vendorNote = sameVendor.length ? null : `Matched by number only (recorded under "${hit.vendor_name}")`;
    if (line.amount != null && hit.amount != null && Math.abs(Number(hit.amount) - line.amount) > 0.009) {
      return {
        ...line,
        recon_status: "amount_mismatch",
        matched_ref: hit.ref,
        note: [`Statement ${line.amount.toFixed(2)} vs recorded ${Number(hit.amount).toFixed(2)}`, vendorNote]
          .filter(Boolean).join("; "),
      };
    }
    return {
      ...line,
      recon_status: ledger ? "found_in_ledger" : "received_by_email",
      matched_ref: hit.ref,
      note: vendorNote,
    };
  });

  return { lines: out, sourceNote };
}
