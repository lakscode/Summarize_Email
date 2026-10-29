// Gmail API access for connected (OAuth) Google mailboxes, mapped onto the shared MailClient shape.
// Gmail has no forward call, so forwarding rebuilds the message (text + original attachments) and sends it.
import { decodeBase64Url, encodeBase64 } from "./base64.ts";
import PostalMime from "npm:postal-mime@3.0.1";
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
import {
  htmlToText, isEmailAttachment, type MailAttachment, type MailboxRow, type MailClient, type MailMessage, readRefreshToken,
} from "./mail.ts";
import { env } from "./supabase.ts";

const GMAIL = "https://gmail.googleapis.com/gmail/v1/users/me";
const UPLOAD = "https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media";

// Requested at connect time: read mail and send (forward invoices, summary email). gmail.readonly is a Google
// "restricted" scope, so a public app needs Google's verification; in Testing mode only listed test users can connect.
export const GOOGLE_SCOPES = "https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/gmail.send";
export const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
export const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";

const FETCH_CONCURRENCY = 5;

interface GmailHeader { name: string; value: string }
interface GmailPart {
  partId?: string;
  mimeType?: string;
  filename?: string;
  headers?: GmailHeader[];
  body?: { size?: number; data?: string; attachmentId?: string };
  parts?: GmailPart[];
}
interface GmailMessage { id: string; threadId: string; internalDate: string; payload: GmailPart }

export async function googleClient(db: SupabaseClient, mailbox: MailboxRow): Promise<MailClient> {
  const token = await googleToken(db, mailbox);
  // Full messages fetched by listInbox, reused for attachments and forwarding.
  const cache = new Map<string, GmailMessage>();
  const full = async (id: string) => {
    let msg = cache.get(id);
    if (!msg) {
      msg = await gmail<GmailMessage>(token, `/messages/${encodeURIComponent(id)}?format=full`);
      cache.set(id, msg);
    }
    return msg;
  };
  const bytes = async (messageId: string, attachmentId: string) => {
    if (attachmentId.startsWith("part:")) {
      const part = flatten((await full(messageId)).payload).find((p) => p.partId === attachmentId.slice(5));
      if (part?.body?.data) return decodeBase64Url(part.body.data);
      // An attached email Gmail expanded into sub-parts: take it from the raw message instead.
      const raw = await gmail<{ raw: string }>(token, `/messages/${encodeURIComponent(messageId)}?format=raw`);
      const parsed = await PostalMime.parse(decodeBase64Url(raw.raw), { forceRfc822Attachments: true });
      const match = parsed.attachments.find((a) => a.filename === part?.filename);
      if (!match) throw new Error(`Attachment ${part?.filename ?? attachmentId} not found in the raw message`);
      return typeof match.content === "string" ? new TextEncoder().encode(match.content) : new Uint8Array(match.content);
    }
    const data = await gmail<{ data: string }>(
      token, `/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`);
    return decodeBase64Url(data.data);
  };

  return {
    async listInbox(since, top = 50) {
      const params = new URLSearchParams({ labelIds: "INBOX", q: `after:${Math.floor(since.getTime() / 1000)}`, maxResults: String(top) });
      const list = await gmail<{ messages?: { id: string }[] }>(token, `/messages?${params}`);
      const ids = (list.messages ?? []).map((m) => m.id);
      const queue = [...ids];
      await Promise.all(Array.from({ length: Math.min(FETCH_CONCURRENCY, queue.length) }, async () => {
        for (let id = queue.shift(); id; id = queue.shift()) await full(id);
      }));
      return ids.map((id) => toMailMessage(cache.get(id)!, mailbox.email_address))
        .sort((a, b) => b.receivedDateTime.localeCompare(a.receivedDateTime));
    },

    async listAttachments(messageId) {
      return attachmentsOf(await full(messageId));
    },

    attachmentBytes: bytes,

    async sendMail(to, subject, html) {
      await send(token, [
        `To: ${to}`,
        `Subject: ${encodeHeader(subject)}`,
        "MIME-Version: 1.0",
        "Content-Type: text/html; charset=UTF-8",
        "Content-Transfer-Encoding: base64",
        "",
        wrap(encodeBase64(html)),
      ].join("\r\n"));
    },

    async forwardMessage(messageId, to, comment) {
      const msg = await full(messageId);
      const original = toMailMessage(msg, mailbox.email_address);
      const header = (name: string) => headerValue(msg.payload, name) ?? "";
      const text = [
        comment,
        "",
        "---------- Forwarded message ---------",
        `From: ${header("From")}`,
        `Date: ${header("Date")}`,
        `Subject: ${header("Subject")}`,
        `To: ${header("To")}`,
        "",
        original.body?.content ?? "",
      ].join("\r\n");

      const boundary = `sm-${crypto.randomUUID()}`;
      const parts = [[
        "Content-Type: text/plain; charset=UTF-8",
        "Content-Transfer-Encoding: base64",
        "",
        wrap(encodeBase64(text)),
      ].join("\r\n")];
      for (const a of attachmentsOf(msg)) {
        parts.push([
          `Content-Type: ${a.contentType || "application/octet-stream"}; ${fileParam("name", a.name)}`,
          `Content-Disposition: attachment; ${fileParam("filename", a.name)}`,
          "Content-Transfer-Encoding: base64",
          "",
          wrap(encodeBase64(await bytes(messageId, a.id))),
        ].join("\r\n"));
      }
      await send(token, [
        `To: ${to}`,
        `Subject: ${encodeHeader(`Fwd: ${original.subject ?? ""}`)}`,
        "MIME-Version: 1.0",
        `Content-Type: multipart/mixed; boundary="${boundary}"`,
        "",
        ...parts.map((p) => `--${boundary}\r\n${p}`),
        `--${boundary}--`,
        "",
      ].join("\r\n"));
    },
  };
}

// Exchanges the stored refresh token for an access token (Google doesn't rotate refresh tokens).
async function googleToken(db: SupabaseClient, mailbox: MailboxRow): Promise<string> {
  const res = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    body: new URLSearchParams({
      client_id: env("GOOGLE_CLIENT_ID"),
      client_secret: env("GOOGLE_CLIENT_SECRET"),
      grant_type: "refresh_token",
      refresh_token: await readRefreshToken(db, mailbox),
    }),
  });
  if (!res.ok) {
    // invalid_grant: access revoked, or the Google app is in Testing mode (tokens there last 7 days).
    throw new Error(`Google token refresh failed (${res.status}); reconnect the mailbox. ${await res.text()}`);
  }
  return (await res.json()).access_token as string;
}

async function gmail<T>(token: string, path: string): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(`${GMAIL}${path}`, { headers: { Authorization: `Bearer ${token}` } });
    if (res.status === 429 && attempt < 3) {
      await new Promise((r) => setTimeout(r, Number(res.headers.get("Retry-After") ?? 2) * 1000));
      continue;
    }
    if (!res.ok) throw new Error(`Gmail GET ${path.split("?")[0]} failed: ${res.status} ${await res.text()}`);
    return (await res.json()) as T;
  }
}

// Sends a raw MIME message (media upload allows messages up to Gmail's 35 MB limit).
async function send(token: string, mime: string): Promise<void> {
  const res = await fetch(UPLOAD, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "message/rfc822" },
    body: mime,
  });
  if (!res.ok) throw new Error(`Gmail send failed: ${res.status} ${await res.text()}`);
}

// --- reading messages -------------------------------------------------------------------------

// The message's parts, not descending into attached emails (their contents are read from the .eml itself).
function flatten(part: GmailPart): GmailPart[] {
  if (part.mimeType === "message/rfc822") return [part];
  return [part, ...(part.parts ?? []).flatMap(flatten)];
}

function headerValue(part: GmailPart, name: string): string | undefined {
  return part.headers?.find((h) => h.name.toLowerCase() === name.toLowerCase())?.value;
}

function attachmentsOf(msg: GmailMessage): MailAttachment[] {
  return flatten(msg.payload).filter((p) => p.filename).map((p) => {
    const disposition = (headerValue(p, "Content-Disposition") ?? "").toLowerCase();
    return {
      id: p.body?.attachmentId ?? `part:${p.partId}`,
      name: p.filename!,
      contentType: p.mimeType,
      size: p.body?.size ?? 0,
      isInline: disposition.startsWith("inline") || (!disposition && !!headerValue(p, "Content-ID")),
      isFile: true,
      isEmail: isEmailAttachment(p.filename!, p.mimeType),
    };
  });
}

function decodeText(part: GmailPart): string {
  const charset = /charset="?([^";]+)"?/i.exec(headerValue(part, "Content-Type") ?? "")?.[1] ?? "utf-8";
  const data = decodeBase64Url(part.body?.data ?? "");
  try {
    return new TextDecoder(charset).decode(data);
  } catch {
    return new TextDecoder().decode(data);
  }
}

// Plain-text body: the text/plain part, or the HTML part with tags stripped.
function bodyText(msg: GmailMessage): string {
  const parts = flatten(msg.payload).filter((p) => !p.filename && p.body?.data);
  const plain = parts.find((p) => p.mimeType === "text/plain");
  if (plain) return decodeText(plain);
  const html = parts.find((p) => p.mimeType === "text/html");
  return html ? htmlToText(decodeText(html)) : "";
}

// "Name <address>" or a bare address.
function parseFrom(value: string | undefined) {
  if (!value) return {};
  const m = /^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/.exec(value);
  return m ? { name: m[1].trim() || undefined, address: m[2].trim() } : { address: value.trim() };
}

function toMailMessage(msg: GmailMessage, mailboxAddress: string): MailMessage {
  return {
    id: msg.id,
    subject: headerValue(msg.payload, "Subject"),
    receivedDateTime: new Date(Number(msg.internalDate)).toISOString(),
    from: { emailAddress: parseFrom(headerValue(msg.payload, "From")) },
    body: { content: bodyText(msg) },
    hasAttachments: attachmentsOf(msg).length > 0,
    webLink: `https://mail.google.com/mail/?authuser=${encodeURIComponent(mailboxAddress)}#all/${msg.id}`,
  };
}

// --- writing MIME -----------------------------------------------------------------------------

const wrap = (base64: string) => base64.replace(/.{1,76}/g, "$&\r\n").trimEnd();

// RFC 2047 encoded words for non-ASCII headers, split so each word stays under the 75-character limit.
function encodeHeader(value: string): string {
  const clean = value.replace(/[\r\n]+/g, " ");
  if (/^[\x20-\x7e]*$/.test(clean)) return clean;
  const words: string[] = [];
  let chunk = "";
  for (const ch of clean) {
    if (new TextEncoder().encode(chunk + ch).length > 45) {
      words.push(chunk);
      chunk = "";
    }
    chunk += ch;
  }
  if (chunk) words.push(chunk);
  return words.map((w) => `=?UTF-8?B?${encodeBase64(w)}?=`).join("\r\n ");
}

// name="..." / filename="..." with an RFC 2231 UTF-8 form for non-ASCII names.
function fileParam(param: string, name: string): string {
  const safe = name.replace(/[\r\n"\\]/g, "_");
  if (/^[\x20-\x7e]*$/.test(safe)) return `${param}="${safe}"`;
  const encoded = encodeURIComponent(name).replace(/['()*!]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
  return `${param}*=UTF-8''${encoded}`;
}
