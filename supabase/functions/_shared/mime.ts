// Minimal RFC 5322 / MIME builder for providers that send raw messages (Gmail).
import type { OutgoingAttachment } from "./types.ts";
import { b64 } from "./util.ts";

function encodeHeader(value: string): string {
  // deno-lint-ignore no-control-regex
  return /^[\x00-\x7F]*$/.test(value) ? value : `=?UTF-8?B?${b64.encode(new TextEncoder().encode(value))}?=`;
}

function wrap76(s: string): string {
  return s.replace(/.{1,76}/g, "$&\r\n").trimEnd();
}

export function buildMime(opts: {
  from: string;
  to: string;
  subject: string;
  html?: string;
  text?: string;
  attachments?: OutgoingAttachment[];
}): string {
  const boundary = `b_${crypto.randomUUID()}`;
  const alt = `a_${crypto.randomUUID()}`;
  const lines: string[] = [
    `From: ${opts.from}`,
    `To: ${opts.to}`,
    `Subject: ${encodeHeader(opts.subject)}`,
    `MIME-Version: 1.0`,
    `Date: ${new Date().toUTCString()}`,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    ``,
    `--${boundary}`,
    `Content-Type: multipart/alternative; boundary="${alt}"`,
    ``,
  ];
  if (opts.text !== undefined) {
    lines.push(`--${alt}`, `Content-Type: text/plain; charset=UTF-8`, `Content-Transfer-Encoding: base64`, ``,
      wrap76(b64.encode(new TextEncoder().encode(opts.text))));
  }
  if (opts.html !== undefined) {
    lines.push(`--${alt}`, `Content-Type: text/html; charset=UTF-8`, `Content-Transfer-Encoding: base64`, ``,
      wrap76(b64.encode(new TextEncoder().encode(opts.html))));
  }
  lines.push(`--${alt}--`);
  for (const a of opts.attachments ?? []) {
    const name = encodeHeader(a.name).replace(/"/g, "");
    lines.push(
      `--${boundary}`,
      `Content-Type: ${a.contentType}; name="${name}"`,
      `Content-Disposition: attachment; filename="${name}"`,
      `Content-Transfer-Encoding: base64`,
      ``,
      wrap76(b64.encode(a.content)),
    );
  }
  lines.push(`--${boundary}--`, ``);
  return lines.join("\r\n");
}
