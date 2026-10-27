// Microsoft Graph access for connected (delegated OAuth) Outlook / Microsoft 365 mailboxes.
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
import {
  isEmailAttachment, type MailAttachment, type MailboxRow, type MailClient, type MailMessage, readRefreshToken, saveRefreshToken,
} from "./mail.ts";
import { env } from "./supabase.ts";

const GRAPH = "https://graph.microsoft.com/v1.0";

// Requested at connect time. Mail.Send is needed to forward invoices to the Smart AP address;
// mailboxes connected before it was added must be reconnected to grant it.
export const MS_SCOPES = "offline_access User.Read Mail.Read Mail.Send";

export const msBase = () => `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0`;

export async function microsoftClient(db: SupabaseClient, mailbox: MailboxRow): Promise<MailClient> {
  const token = await graphToken(db, mailbox);
  return {
    listInbox: (since, top) => listInbox(token, since, top),
    listAttachments: (messageId) => listAttachments(token, messageId),
    attachmentBytes: (messageId, attachmentId) => attachmentBytes(token, messageId, attachmentId),
    sendMail: (to, subject, html) => sendMail(token, to, subject, html),
    forwardMessage: (messageId, to, comment) => forwardMessage(token, messageId, to, comment),
  };
}

// Exchanges the stored refresh token for an access token, saving a rotated refresh token if Microsoft returns one.
async function graphToken(db: SupabaseClient, mailbox: MailboxRow): Promise<string> {
  const refreshToken = await readRefreshToken(db, mailbox);
  const res = await fetch(`${msBase()}/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      // .default = whatever this mailbox consented to, so older connections without Mail.Send still refresh.
      scope: "offline_access https://graph.microsoft.com/.default",
    }),
  });
  if (!res.ok) throw new Error(`Microsoft token refresh failed (${res.status}); reconnect the mailbox. ${await res.text()}`);
  const tok = await res.json();
  if (tok.refresh_token && tok.refresh_token !== refreshToken) await saveRefreshToken(db, mailbox, tok.refresh_token);
  return tok.access_token as string;
}

async function graph<T>(token: string, path: string, init: RequestInit = {}, extraHeaders: Record<string, string> = {}): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
      ...init,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...extraHeaders },
    });
    if (res.status === 429 && attempt < 3) {
      await new Promise((r) => setTimeout(r, Number(res.headers.get("Retry-After") ?? 2) * 1000));
      continue;
    }
    if (!res.ok) throw new Error(`Graph ${init.method ?? "GET"} ${path.split("?")[0]} failed: ${res.status} ${await res.text()}`);
    if (res.status === 202 || res.status === 204) return undefined as T;
    return (await res.json()) as T;
  }
}

// Inbox messages received since `since`, newest first, bodies as plain text.
async function listInbox(token: string, since: Date, top = 50): Promise<MailMessage[]> {
  const params = new URLSearchParams({
    $filter: `receivedDateTime ge ${since.toISOString()}`,
    $orderby: "receivedDateTime desc",
    $top: String(top),
    $select: "id,internetMessageId,subject,receivedDateTime,from,body,hasAttachments,webLink",
  });
  const data = await graph<{ value: MailMessage[] }>(token, `/me/mailFolders/inbox/messages?${params}`, {}, {
    Prefer: 'outlook.body-content-type="text"',
  });
  return data.value;
}

interface GraphAttachment {
  id: string;
  name: string;
  contentType?: string;
  size: number;
  isInline?: boolean;
  "@odata.type"?: string;
}

async function listAttachments(token: string, messageId: string): Promise<MailAttachment[]> {
  const data = await graph<{ value: GraphAttachment[] }>(
    token,
    `/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline`,
  );
  return data.value.map((a) => {
    const isFile = !a["@odata.type"] || a["@odata.type"] === "#microsoft.graph.fileAttachment";
    // Outlook items (e.g. "attach as item" forwards) download as MIME through $value, like .eml files.
    const isItem = a["@odata.type"] === "#microsoft.graph.itemAttachment";
    return {
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
      isInline: a.isInline,
      isFile,
      isEmail: isItem || (isFile && isEmailAttachment(a.name, a.contentType)),
    };
  });
}

// Raw bytes of a file attachment, or the MIME of an attached Outlook item.
async function attachmentBytes(token: string, messageId: string, attachmentId: string): Promise<Uint8Array> {
  const res = await fetch(
    `${GRAPH}/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (!res.ok) throw new Error(`Graph GET attachment failed: ${res.status} ${await res.text()}`);
  return new Uint8Array(await res.arrayBuffer());
}

// Sends an HTML email from the connected mailbox (saved to its Sent Items).
async function sendMail(token: string, to: string, subject: string, html: string): Promise<void> {
  await graph<void>(token, "/me/sendMail", {
    method: "POST",
    body: JSON.stringify({
      message: { subject, body: { contentType: "HTML", content: html }, toRecipients: [{ emailAddress: { address: to } }] },
      saveToSentItems: true,
    }),
  });
}

async function forwardMessage(token: string, messageId: string, to: string, comment: string): Promise<void> {
  await graph<void>(token, `/me/messages/${encodeURIComponent(messageId)}/forward`, {
    method: "POST",
    body: JSON.stringify({ comment, toRecipients: [{ emailAddress: { address: to } }] }),
  });
}
