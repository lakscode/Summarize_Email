// Deterministic safety signals (from the ap-email-triage fraud checks). These are passed to the
// classifier as evidence and also enforced in code, so a flagged email is never auto-forwarded
// even if the model under-calls the risk.
import type { MailMessage } from "./types.ts";
import { domainOf, registrableDomain } from "./util.ts";

export interface SafetyReport {
  strong: string[]; // any one of these flags the email
  weak: string[];   // flags only together with another signal
}

const FREE_MAIL = new Set([
  "gmail.com", "googlemail.com", "yahoo.com", "outlook.com", "hotmail.com", "live.com", "aol.com",
  "icloud.com", "proton.me", "protonmail.com", "gmx.com", "mail.com", "yandex.com", "zoho.com",
]);

const RISKY_EXT = /\.(zip|rar|7z|html?|exe|js|vbs|scr|bat|cmd|iso|img|docm|xlsm|pptm|one|lnk)$/i;

const BANK_CHANGE =
  /(new|updated|changed?|change of)\s+(bank(ing)?|remittance|payment|account|wire|ach)\s+(details|information|info|instructions|address)|remit(tance)?\s+to\s+(our\s+)?new|please\s+update\s+(our|your)\s+(records|vendor|bank)|confirm\s+in\s+your\s+system/i;

const PRESSURE = /\b(urgent(ly)?\s+(wire|payment|transfer)|gift\s*cards?|pay\s+(today|immediately|within\s+24)|wire\s+(today|immediately))\b/i;

export function safetyCheck(msg: MailMessage): SafetyReport {
  const strong: string[] = [];
  const weak: string[] = [];
  const fromDomain = registrableDomain(domainOf(msg.from.address));
  const text = `${msg.subject}\n${msg.bodyText}`;

  if (BANK_CHANGE.test(text)) strong.push("Asks to change or confirm bank / remittance details");
  if (PRESSURE.test(text)) strong.push("Pressure for unusually fast payment (wire / gift cards / pay today)");

  for (const a of msg.attachments) {
    if (RISKY_EXT.test(a.name)) strong.push(`Unusual attachment type for an invoice: ${a.name}`);
  }

  const replyDomains = msg.replyTo.map((r) => registrableDomain(domainOf(r))).filter((d) => d && d !== fromDomain);
  if (replyDomains.length) weak.push(`Reply-To domain (${replyDomains.join(", ")}) differs from sender (${fromDomain})`);

  if (FREE_MAIL.has(fromDomain)) weak.push(`Sent from a free email address (${fromDomain})`);

  const auth = (msg.headers["authentication-results"] ?? "").toLowerCase();
  for (const mech of ["spf", "dkim", "dmarc"]) {
    const m = auth.match(new RegExp(`${mech}=(\\w+)`));
    if (m && ["fail", "softfail", "permerror"].includes(m[1])) strong.push(`${mech.toUpperCase()} ${m[1]}`);
    else if (m && ["none", "neutral"].includes(m[1])) weak.push(`${mech.toUpperCase()} ${m[1]}`);
  }

  return { strong, weak };
}

export function isFlagged(report: SafetyReport, modelRisk: "none" | "low" | "high"): boolean {
  return report.strong.length > 0 || report.weak.length >= 2 || modelRisk === "high";
}

/** Only keep a portal link when it points at the sender's own domain or a configured portal. */
export function trustedPortalUrl(
  url: string | null,
  senderAddress: string,
  portalVendors: { name: string; portal: string }[],
): { url: string | null; warning: string | null } {
  if (!url) return { url: null, warning: null };
  let host: string;
  try {
    host = new URL(url).hostname;
  } catch {
    return { url: null, warning: "Portal link in the email is malformed - go to the provider's website directly." };
  }
  const linkDomain = registrableDomain(host);
  const allowed = new Set([registrableDomain(domainOf(senderAddress))]);
  for (const p of portalVendors) {
    try {
      allowed.add(registrableDomain(new URL(p.portal).hostname));
    } catch { /* ignore bad config */ }
  }
  if (allowed.has(linkDomain)) return { url, warning: null };
  return {
    url: null,
    warning: `Portal link points to ${linkDomain}, not the provider's own site - type the provider's address yourself.`,
  };
}
