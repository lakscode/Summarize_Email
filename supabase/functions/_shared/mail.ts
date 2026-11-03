// Provider-neutral mailbox access used by process.ts. Messages use the Microsoft Graph shape (the first
// provider); the Gmail adapter maps its messages into it.
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
import { encodeBase64 } from "./base64.ts";
import { googleClient } from "./google.ts";
import { microsoftClient } from "./microsoft.ts";

export type Provider = "microsoft" | "google";
export const PROVIDERS: Provider[] = ["microsoft", "google"];
export const PROVIDER_LABELS: Record<Provider, string> = { microsoft: "Microsoft", google: "Google" };

export interface MailboxRow {
  id: string;
  provider: Provider;
  email_address: string;
  credentials_secret_id: string | null;
  // Microsoft shared mailbox: the signed-in account whose token reads `email_address`. Null for an own mailbox.
  delegate_address?: string | null;
  // "app": Microsoft API keys (client credentials) of the organization connection `connection_id`; the mailbox has
  // no credentials of its own.
  auth_type?: "oauth" | "app";
  connection_id?: string | null;
}

export interface MailMessage {
  id: string;
  subject?: string;
  receivedDateTime: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  body?: { content?: string }; // plain text
  hasAttachments?: boolean;
  webLink?: string;
}

export interface MailAttachment {
  id: string;
  name: string;
  contentType?: string;
  size: number;
  isInline?: boolean;
  // False for things that aren't files, such as Outlook item or reference attachments.
  isFile: boolean;
  // An attached email (.eml / Outlook item); attachmentBytes returns its raw MIME.
  isEmail?: boolean;
}

// One message as the Inbox preview shows it: the HTML body when there is one (otherwise text), headers as display
// strings, and its images referenced by cid: (already loaded, within MAX_INLINE_IMAGE_BYTES).
export interface FullMessage {
  subject?: string;
  from?: string;
  to?: string;
  cc?: string;
  received: string;
  html: string | null;
  text: string;
  attachments: MailAttachment[];
  inlineImages: InlineImage[];
  webLink?: string;
}
export interface InlineImage { contentId: string; contentType: string; bytes: Uint8Array }

// Total size of the cid: images loaded for one preview; bigger ones show as broken images.
export const MAX_INLINE_IMAGE_BYTES = 5 * 1024 * 1024;
export const contentIdOf = (value: string | undefined) => value?.trim().replace(/^<|>$/g, "") || null;

export const isEmailAttachment = (name: string, contentType?: string) =>
  (contentType ?? "").toLowerCase() === "message/rfc822" || name.toLowerCase().endsWith(".eml");

// Readable text from an HTML email body or attachment.
export function htmlToText(html: string): string {
  return html
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>|<\/(p|div|tr|li|h\d)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export interface MailClient {
  // Ids of the inbox messages received since `since`, newest first, at most `max` (cheap: no contents).
  listInboxIds(since: Date, max: number): Promise<string[]>;
  // Those messages, in the same order; ones no longer in the mailbox (moved, deleted) are left out.
  getMessages(ids: string[]): Promise<MailMessage[]>;
  listAttachments(messageId: string): Promise<MailAttachment[]>;
  // The whole message for the Inbox preview.
  getMessage(messageId: string): Promise<FullMessage>;
  attachmentBytes(messageId: string, attachmentId: string): Promise<Uint8Array>;
  // Sends an HTML email to one or more addresses from the connected mailbox (kept in its sent mail).
  sendMail(to: string[], subject: string, html: string): Promise<void>;
  // Forwards the message with its attachments, `comment` above the original.
  forwardMessage(messageId: string, to: string, comment: string): Promise<void>;
  // Of these messages, the ones someone has forwarded (e.g. manually in Outlook), with when. Only where the provider
  // records it on the message (Microsoft 365); messages no longer in the mailbox are left out.
  forwardedAt?(messageIds: string[]): Promise<Map<string, string>>;
}

// fn over items, at most `limit` at a time; results in the items' order.
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results = new Array<R>(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    for (let i = next++; i < items.length; i = next++) results[i] = await fn(items[i]);
  }));
  return results;
}

// Whether a provider error means the message is gone (deleted, or moved: Outlook gives a moved message a new id).
export const isNotFound = (err: unknown) => /failed: 404\b/.test(String(err));

// Replaces cid: references in an HTML body with data: URLs of the message's inline images.
export function embedInlineImages(html: string, images: InlineImage[]): string {
  const byId = new Map(images.map((i) => [i.contentId.toLowerCase(), i]));
  return html.replace(/cid:([^"'\s)>]+)/gi, (match, id: string) => {
    const image = byId.get(decodeURIComponent(id).toLowerCase());
    return image ? `data:${image.contentType};base64,${encodeBase64(image.bytes)}` : match;
  });
}

export async function mailClient(db: SupabaseClient, mailbox: MailboxRow): Promise<MailClient> {
  return mailbox.provider === "google" ? await googleClient(db, mailbox) : await microsoftClient(db, mailbox);
}

// The mailbox's refresh token from Vault.
export async function readRefreshToken(db: SupabaseClient, mailbox: MailboxRow): Promise<string> {
  if (!mailbox.credentials_secret_id) throw new Error("Mailbox has no stored credentials; reconnect it.");
  const { data, error } = await db.rpc("mailbox_secret_get", { p_secret_id: mailbox.credentials_secret_id });
  if (error || !data) throw new Error(`Could not read mailbox credentials: ${error?.message ?? "missing"}`);
  return data as string;
}

// Stores a rotated refresh token; failures are logged, since the old token usually keeps working.
export async function saveRefreshToken(db: SupabaseClient, mailbox: MailboxRow, token: string): Promise<void> {
  const { error } = await db.rpc("mailbox_secret_put", {
    p_secret_id: mailbox.credentials_secret_id,
    p_secret: token,
    p_name: `mailbox:${mailbox.id}`,
  });
  if (error) console.error("Could not save rotated refresh token:", error.message);
}
