// Provider-neutral mailbox access used by process.ts. Messages use the Microsoft Graph shape (the first
// provider); the Gmail adapter maps its messages into it.
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
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
  // "app": Microsoft API keys (client credentials); the Vault secret is AppCredentials JSON, not a refresh token.
  auth_type?: "oauth" | "app";
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
  // Inbox messages received since `since`, newest first.
  listInbox(since: Date, top?: number): Promise<MailMessage[]>;
  listAttachments(messageId: string): Promise<MailAttachment[]>;
  attachmentBytes(messageId: string, attachmentId: string): Promise<Uint8Array>;
  // Sends an HTML email from the connected mailbox (kept in its sent mail).
  sendMail(to: string, subject: string, html: string): Promise<void>;
  // Forwards the message with its attachments, `comment` above the original.
  forwardMessage(messageId: string, to: string, comment: string): Promise<void>;
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
