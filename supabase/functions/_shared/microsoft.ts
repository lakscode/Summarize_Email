// Microsoft Graph access for connected (delegated OAuth) mailboxes.
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
import { env } from "./supabase.ts";

const GRAPH = "https://graph.microsoft.com/v1.0";

// Requested at connect time. Mail.Send is needed to forward invoices to the Smart AP address;
// mailboxes connected before it was added must be reconnected to grant it.
export const MS_SCOPES = "offline_access User.Read Mail.Read Mail.Send";

export const msBase = () => `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0`;

export interface MailboxRow {
  id: string;
  email_address: string;
  credentials_secret_id: string | null;
}

export interface GraphMessage {
  id: string;
  internetMessageId?: string;
  subject?: string;
  receivedDateTime: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  body?: { content?: string };
  hasAttachments?: boolean;
  webLink?: string;
}

// Exchanges the stored refresh token for an access token, saving a rotated refresh token if Microsoft returns one.
export async function graphToken(db: SupabaseClient, mailbox: MailboxRow): Promise<string> {
  if (!mailbox.credentials_secret_id) throw new Error("Mailbox has no stored credentials; reconnect it.");
  const { data: refreshToken, error } = await db.rpc("mailbox_secret_get", { p_secret_id: mailbox.credentials_secret_id });
  if (error || !refreshToken) throw new Error(`Could not read mailbox credentials: ${error?.message ?? "missing"}`);

  const res = await fetch(`${msBase()}/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
      grant_type: "refresh_token",
      refresh_token: refreshToken as string,
      // .default = whatever this mailbox consented to, so older connections without Mail.Send still refresh.
      scope: "offline_access https://graph.microsoft.com/.default",
    }),
  });
  if (!res.ok) throw new Error(`Microsoft token refresh failed (${res.status}); reconnect the mailbox. ${await res.text()}`);
  const tok = await res.json();

  if (tok.refresh_token && tok.refresh_token !== refreshToken) {
    const { error: putError } = await db.rpc("mailbox_secret_put", {
      p_secret_id: mailbox.credentials_secret_id,
      p_secret: tok.refresh_token,
      p_name: `mailbox:${mailbox.id}`,
    });
    if (putError) console.error("Could not save rotated refresh token:", putError.message);
  }
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
export async function listInbox(token: string, since: Date, top = 50): Promise<GraphMessage[]> {
  const params = new URLSearchParams({
    $filter: `receivedDateTime ge ${since.toISOString()}`,
    $orderby: "receivedDateTime desc",
    $top: String(top),
    $select: "id,internetMessageId,subject,receivedDateTime,from,body,hasAttachments,webLink",
  });
  const data = await graph<{ value: GraphMessage[] }>(token, `/me/mailFolders/inbox/messages?${params}`, {}, {
    Prefer: 'outlook.body-content-type="text"',
  });
  return data.value;
}

export async function attachmentNames(token: string, messageId: string): Promise<string[]> {
  const data = await graph<{ value: { name: string }[] }>(
    token,
    `/me/messages/${encodeURIComponent(messageId)}/attachments?$select=name`,
  );
  return data.value.map((a) => a.name);
}

export async function forwardMessage(token: string, messageId: string, to: string, comment: string): Promise<void> {
  await graph<void>(token, `/me/messages/${encodeURIComponent(messageId)}/forward`, {
    method: "POST",
    body: JSON.stringify({ comment, toRecipients: [{ emailAddress: { address: to } }] }),
  });
}
