// Live check of the Claude classifier on sample emails (no mailbox or database needed).
//   ANTHROPIC_API_KEY=... deno run -A supabase/functions/tests/classify_smoke.ts
import { classifyEmail } from "../_shared/classify.ts";
import { safetyCheck } from "../_shared/safety.ts";
import type { MailMessage, MailProvider } from "../_shared/types.ts";

const fake = { getAttachment: () => Promise.resolve(new Uint8Array()) } as unknown as MailProvider;

const base = {
  id: "x", threadId: null, replyTo: [], to: ["ap@example.com"], receivedAt: new Date(), attachments: [], labels: [],
  headers: { "authentication-results": "spf=pass dkim=pass dmarc=pass" },
};

const samples: MailMessage[] = [
  {
    ...base,
    from: { address: "billing@titansecurity.com", name: "Titan Security" },
    subject: "Invoice 483739",
    bodyText: "Invoice 483739\nDate: 2026-10-20\nBill to: Park Madison LLC\nGuard services October\nAmount due: $1,240.00\nDue: 2026-11-15",
  },
  {
    ...base,
    from: { address: "noreply@evergy.com", name: "Evergy" },
    subject: "Your Evergy bill is ready",
    bodyText: "Your bill for account ending 4471 is ready. Amount due $312.18 by 2026-11-03. View and pay at https://www.evergy.com/account",
  },
  {
    ...base,
    from: { address: "ar@cookspest.com", name: "Cook's Pest Control" },
    subject: "Statement of account",
    bodyText: "Statement 2026-10-01\nInvoice 29142719  2026-09-02  $425.00 open\nInvoice 29305880  2026-09-16  $425.00 open\nBalance due $850.00",
  },
];

for (const msg of samples) {
  const { classification: c } = await classifyEmail(fake, msg, {
    portalVendors: [{ name: "Evergy", portal: "https://www.evergy.com/" }],
    managedEntities: [],
    safety: safetyCheck(msg),
  });
  console.log(`\n${msg.subject}\n  -> ${c.category} (${c.confidence}), fraud ${c.fraud_risk}\n  ${c.employee_action}`);
  if (c.statement) console.log("  lines:", c.statement.lines.map((l) => `${l.invoice_number} ${l.amount}`).join(", "));
}
