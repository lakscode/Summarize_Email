import { assert, assertEquals, assertStringIncludes } from "jsr:@std/assert@1.0.15";
import { isFlagged, safetyCheck, trustedPortalUrl } from "../_shared/safety.ts";
import { buildMime } from "../_shared/mime.ts";
import type { MailMessage } from "../_shared/types.ts";
import { normalizeInvoiceNumber, registrableDomain } from "../_shared/util.ts";

function msg(overrides: Partial<MailMessage>): MailMessage {
  return {
    id: "1",
    threadId: null,
    from: { address: "billing@titansecurity.com", name: "Titan Security" },
    replyTo: [],
    to: ["ap@example.com"],
    subject: "Invoice 483739",
    receivedAt: new Date("2026-10-20T10:00:00Z"),
    bodyText: "Please find attached invoice 483739 for $1,240.00 due 15 Nov.",
    headers: { "authentication-results": "spf=pass dkim=pass dmarc=pass" },
    attachments: [{ id: "a", name: "INV-483739.pdf", contentType: "application/pdf", size: 1000 }],
    labels: [],
    ...overrides,
  };
}

Deno.test("clean invoice is not flagged", () => {
  const r = safetyCheck(msg({}));
  assertEquals(r.strong, []);
  assertEquals(r.weak, []);
  assert(!isFlagged(r, "none"));
});

Deno.test("remittance change is a strong signal", () => {
  const r = safetyCheck(msg({
    bodyText: "Our remittance address has changed. Please update our records: remit to new PO Box 95790, Chicago.",
  }));
  assert(r.strong.some((s) => s.includes("bank / remittance")));
  assert(isFlagged(r, "none"));
});

Deno.test("DKIM none alone is not enough; with a reply-to mismatch it is", () => {
  const dkimOnly = safetyCheck(msg({ headers: { "authentication-results": "spf=pass dkim=none dmarc=pass" } }));
  assert(!isFlagged(dkimOnly, "none"));
  const both = safetyCheck(msg({
    headers: { "authentication-results": "spf=pass dkim=none dmarc=pass" },
    replyTo: ["payments@titan-security-billing.net"],
  }));
  assert(isFlagged(both, "none"));
});

Deno.test("risky attachment types are flagged", () => {
  const r = safetyCheck(msg({ attachments: [{ id: "a", name: "invoice.html", contentType: "text/html", size: 10 }] }));
  assert(isFlagged(r, "none"));
});

Deno.test("model high risk flags even without deterministic signals", () => {
  assert(isFlagged({ strong: [], weak: [] }, "high"));
});

Deno.test("portal link must match sender or configured portal", () => {
  const vendors = [{ name: "Evergy", portal: "https://www.evergy.com/" }];
  assertEquals(trustedPortalUrl("https://account.evergy.com/login", "noreply@evergy.com", vendors).url,
    "https://account.evergy.com/login");
  assertEquals(trustedPortalUrl("https://www.evergy.com/pay", "alerts@mailer.example.net", vendors).url,
    "https://www.evergy.com/pay");
  const bad = trustedPortalUrl("https://evergy-billing.co/login", "noreply@evergy.com", vendors);
  assertEquals(bad.url, null);
  assertStringIncludes(bad.warning!, "evergy-billing.co");
});

Deno.test("domain and invoice number helpers", () => {
  assertEquals(registrableDomain("billing.acme.co.uk"), "acme.co.uk");
  assertEquals(registrableDomain("mail.acme.com"), "acme.com");
  assertEquals(normalizeInvoiceNumber("inv-20 401"), "INV20401");
});

Deno.test("MIME message carries attachment and encoded subject", () => {
  const mime = buildMime({
    from: "ap@example.com",
    to: "smartap@example.com",
    subject: "Fwd: Invoice – 483739",
    text: "hello",
    attachments: [{ name: "a.pdf", contentType: "application/pdf", content: new Uint8Array([37, 80, 68, 70]) }],
  });
  assertStringIncludes(mime, "Subject: =?UTF-8?B?");
  assertStringIncludes(mime, 'Content-Disposition: attachment; filename="a.pdf"');
  assertStringIncludes(mime, "JVBERg=="); // "%PDF"
});
