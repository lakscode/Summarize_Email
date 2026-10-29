// Microsoft Graph access for connected Outlook / Microsoft 365 mailboxes: delegated OAuth (a user signed in), or
// API keys (app-only client credentials from the customer's own Azure app registration).
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";
import {
  isEmailAttachment, type MailAttachment, type MailboxRow, type MailClient, type MailMessage, readRefreshToken, saveRefreshToken,
} from "./mail.ts";
import { env } from "./supabase.ts";

const GRAPH = "https://graph.microsoft.com/v1.0";

// Requested at connect time. Mail.Send is needed to forward invoices to the Smart AP address;
// mailboxes connected before it was added must be reconnected to grant it.
export const MS_SCOPES = "offline_access User.Read Mail.Read Mail.Send";
// Requested on every Microsoft sign-in, so shared mailboxes (chosen up front or found afterwards) can be read and
// forwarded from with the same token (forwarding also needs Send As or Send on Behalf rights in Exchange).
export const MS_SHARED_SCOPES = `${MS_SCOPES} Mail.Read.Shared Mail.Send.Shared`;

// Graph path of a mailbox: the signed-in user's own, or a shared mailbox they have access to.
export const mailboxPath = (shared?: string | null) => shared ? `/users/${encodeURIComponent(shared)}` : "/me";

export const msBase = () => `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0`;

// Customer's Azure app registration for an organization connection (ms_connections); needs application permissions
// Mail.Read and Mail.Send (admin-consented). Stored as JSON in Vault.
export interface AppCredentials { tenant_id: string; client_id: string; client_secret: string }

export async function connectionCredentials(db: SupabaseClient, connectionId: string): Promise<AppCredentials> {
  const { data, error } = await db.from("ms_connections").select("credentials_secret_id").eq("id", connectionId).maybeSingle();
  if (error) throw new Error(error.message);
  if (!data?.credentials_secret_id) throw new Error("The organization connection for this mailbox has no API keys; update them on the dashboard.");
  const { data: secret, error: secretError } = await db.rpc("mailbox_secret_get", { p_secret_id: data.credentials_secret_id });
  if (secretError || !secret) throw new Error(`Could not read the organization's API keys: ${secretError?.message ?? "missing"}`);
  return JSON.parse(secret as string);
}

export async function microsoftClient(db: SupabaseClient, mailbox: MailboxRow): Promise<MailClient> {
  const app = mailbox.auth_type === "app";
  if (app && !mailbox.connection_id) throw new Error("Mailbox has no organization connection; add it again from the dashboard.");
  const token = app ? await appToken(await connectionCredentials(db, mailbox.connection_id!)) : await delegatedToken(db, mailbox);
  // App-only tokens have no "me"; they address the mailbox by its address.
  const box = mailboxPath(app || mailbox.delegate_address ? mailbox.email_address : null);
  return {
    listInbox: (since, top) => listInbox(token, box, since, top),
    listAttachments: (messageId) => listAttachments(token, box, messageId),
    attachmentBytes: (messageId, attachmentId) => attachmentBytes(token, box, messageId, attachmentId),
    // Delegated: always from the signed-in user's own mailbox, so a shared mailbox's summary needs no Send As rights.
    sendMail: (to, subject, html) => sendMail(token, app ? box : "/me", to, subject, html),
    forwardMessage: (messageId, to, comment) => forwardMessage(token, box, messageId, to, comment),
  };
}

// Exchanges the stored refresh token for an access token, saving a rotated refresh token if Microsoft returns one.
// `resource`: Graph by default, or Exchange Online for Autodiscover (alternateMailboxes).
export async function delegatedToken(db: SupabaseClient, mailbox: MailboxRow, resource = "https://graph.microsoft.com"): Promise<string> {
  const refreshToken = await readRefreshToken(db, mailbox);
  const res = await fetch(`${msBase()}/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      // .default = whatever this mailbox consented to, so older connections without Mail.Send still refresh.
      scope: `offline_access ${resource}/.default`,
    }),
  });
  if (!res.ok) throw new Error(`Microsoft token refresh failed (${res.status}); reconnect the mailbox. ${await res.text()}`);
  const tok = await res.json();
  if (tok.refresh_token && tok.refresh_token !== refreshToken) await saveRefreshToken(db, mailbox, tok.refresh_token);
  return tok.access_token as string;
}

const EXCHANGE = "https://outlook.office365.com";

// Client-credentials access token for an API-key mailbox (Graph), or for the Exchange Online admin API.
export async function appToken(c: AppCredentials, resource = "https://graph.microsoft.com"): Promise<string> {
  const res = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(c.tenant_id)}/oauth2/v2.0/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: c.client_id,
      client_secret: c.client_secret,
      grant_type: "client_credentials",
      scope: `${resource}/.default`,
    }),
  });
  if (!res.ok) {
    const detail = await res.json().catch(() => ({}));
    throw new Error(`Microsoft rejected the API keys (${res.status}${detail.error ? `: ${detail.error}` : ""}); check the tenant ID, client ID and secret.`);
  }
  return (await res.json()).access_token as string;
}

// --- Exchange Online admin API: which shared mailboxes a person has Full Access to ------------------------------
// Graph can't read mailbox permissions, so this runs Exchange cmdlets through the admin API the
// ExchangeOnlineManagement module uses. The Azure app needs the application permission Exchange.ManageAsApp
// (Office 365 Exchange Online) with admin consent, and an Entra role that can read recipients and their
// permissions (Global Reader, or Exchange Recipient Administrator).

// At most this many shared mailboxes are checked per search; each needs one permission lookup.
const MAX_SHARED_TO_CHECK = 500;
const PERMISSION_CONCURRENCY = 8;

export interface SharedMailbox { address: string; name: string }

async function exoCommand<T>(token: string, tenantId: string, cmdlet: string, parameters: Record<string, unknown>): Promise<T[]> {
  const body = JSON.stringify({ CmdletInput: { CmdletName: cmdlet, Parameters: parameters } });
  const out: T[] = [];
  let url: string | undefined = `${EXCHANGE}/adminapi/beta/${encodeURIComponent(tenantId)}/InvokeCommand`;
  for (let page = 0; url && page < 50; page++) {
    const res: Response = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body,
    });
    if (!res.ok) {
      const text = (await res.text()).slice(0, 500);
      if (res.status === 401 || res.status === 403) {
        throw new Error("The Azure app can't run Exchange commands. Give it the Exchange.ManageAsApp application permission " +
          "(Office 365 Exchange Online, admin-consented) and the Global Reader or Exchange Recipient Administrator role.");
      }
      throw new Error(`Exchange ${cmdlet} failed (${res.status}): ${text}`);
    }
    const data: { value?: T[]; "@odata.nextLink"?: string } = await res.json();
    out.push(...(data.value ?? []));
    url = data["@odata.nextLink"];
  }
  return out;
}

// Shared mailboxes in the tenant on which `person` has a direct (not group-inherited), allowed FullAccess entry.
export async function sharedMailboxesFor(c: AppCredentials, person: string): Promise<{ mailboxes: SharedMailbox[]; checked: number; total: number }> {
  const token = await appToken(c, EXCHANGE);
  const shared = await exoCommand<{ PrimarySmtpAddress?: string; DisplayName?: string }>(token, c.tenant_id, "Get-Mailbox", {
    RecipientTypeDetails: "SharedMailbox",
    ResultSize: "Unlimited",
  });
  const candidates = shared.filter((m) => m.PrimarySmtpAddress).slice(0, MAX_SHARED_TO_CHECK);

  const found: SharedMailbox[] = [];
  const queue = [...candidates];
  await Promise.all(Array.from({ length: Math.min(PERMISSION_CONCURRENCY, queue.length) }, async () => {
    for (let m = queue.shift(); m; m = queue.shift()) {
      const entries = await exoCommand<{ AccessRights?: string[] | string; Deny?: boolean | string }>(
        token, c.tenant_id, "Get-MailboxPermission", { Identity: m.PrimarySmtpAddress, User: person },
      ).catch(() => []); // a mailbox that can't be read is skipped rather than failing the whole search
      const full = entries.some((e) => String(e.AccessRights ?? "").includes("FullAccess") && String(e.Deny) !== "true");
      if (full) found.push({ address: m.PrimarySmtpAddress!.toLowerCase(), name: m.DisplayName ?? "" });
    }
  }));
  found.sort((a, b) => a.address.localeCompare(b.address));
  return { mailboxes: found, checked: candidates.length, total: shared.length };
}

// --- Autodiscover: which shared mailboxes a signed-in user has Full Access to ------------------------------------
// Exchange lists mailboxes the user has Full Access to (with automapping, the default when access is granted to the
// user directly) as "Delegate" AlternativeMailbox entries in the Outlook Autodiscover response. Needs a delegated
// token for Exchange Online: the Azure app needs the delegated permission EWS.AccessAsUser.All
// (Office 365 Exchange Online), admin-consented.
// `entries` counts every AlternativeMailbox (any type) and `response` is the start of Exchange's reply, so an empty
// result can be told apart from an unexpected response in the logs.
export async function alternateMailboxes(
  db: SupabaseClient, mailbox: MailboxRow,
): Promise<{ mailboxes: SharedMailbox[]; entries: number; response: string }> {
  const token = await delegatedToken(db, mailbox, EXCHANGE);
  let address = mailbox.email_address;
  let xml = "";
  // Exchange may answer with a redirect to the user's primary address (e.g. when signed in by an alias or UPN).
  for (let hop = 0; hop < 3; hop++) {
    const res = await fetch(`${EXCHANGE}/autodiscover/autodiscover.xml`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/xml; charset=utf-8", "X-AnchorMailbox": address },
      body: `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
  <Request>
    <EMailAddress>${xmlEscape(address)}</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
  </Request>
</Autodiscover>`,
    });
    xml = await res.text();
    if (!res.ok) throw new Error(`Autodiscover failed (${res.status}): ${xml.slice(0, 500)}`);
    const redirect = xmlField(xml, "Action") === "redirectAddr" ? xmlField(xml, "RedirectAddr").toLowerCase() : "";
    if (!redirect || redirect === address) break;
    address = redirect;
  }
  if (/<(\w+:)?Error[\s>]/.test(xml)) {
    throw new Error(`Autodiscover returned an error: ${xmlField(xml, "Message") || xmlField(xml, "ErrorCode") || xml.slice(0, 500)}`);
  }

  const found = new Map<string, SharedMailbox>();
  const entries = [...xml.matchAll(/<(?:\w+:)?AlternativeMailbox>([\s\S]*?)<\/(?:\w+:)?AlternativeMailbox>/g)];
  for (const [, entry] of entries) {
    if (xmlField(entry, "Type") !== "Delegate") continue;
    const smtp = (xmlField(entry, "SmtpAddress") || xmlField(entry, "OwnerSmtpAddress")).toLowerCase();
    if (smtp && smtp !== mailbox.email_address && smtp !== address) found.set(smtp, { address: smtp, name: xmlField(entry, "DisplayName") });
  }
  return {
    mailboxes: [...found.values()].sort((a, b) => a.address.localeCompare(b.address)),
    entries: entries.length,
    response: xml.slice(0, 3000),
  };
}

// Text of the first <name> element (namespace prefix allowed), unescaped.
const xmlField = (xml: string, name: string) =>
  xmlUnescape(xml.match(new RegExp(`<(?:\\w+:)?${name}>([^<]*)</(?:\\w+:)?${name}>`))?.[1] ?? "").trim();

const xmlEscape = (s: string) => s.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
const xmlUnescape = (s: string) =>
  s.replace(/&(lt|gt|quot|apos|amp|#\d+);/g, (_, e) =>
    ({ lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" } as Record<string, string>)[e] ?? String.fromCharCode(Number(e.slice(1))));

// Whether `token` can open the inbox of `address` (the Graph response, for its status and error text).
export const openInbox = (token: string, address: string) =>
  fetch(`${GRAPH}${mailboxPath(address)}/mailFolders/inbox?$select=id`, { headers: { Authorization: `Bearer ${token}` } });

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
async function listInbox(token: string, box: string, since: Date, top = 50): Promise<MailMessage[]> {
  const params = new URLSearchParams({
    $filter: `receivedDateTime ge ${since.toISOString()}`,
    $orderby: "receivedDateTime desc",
    $top: String(top),
    $select: "id,internetMessageId,subject,receivedDateTime,from,body,hasAttachments,webLink",
  });
  const data = await graph<{ value: MailMessage[] }>(token, `${box}/mailFolders/inbox/messages?${params}`, {}, {
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

async function listAttachments(token: string, box: string, messageId: string): Promise<MailAttachment[]> {
  const data = await graph<{ value: GraphAttachment[] }>(
    token,
    `${box}/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline`,
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
async function attachmentBytes(token: string, box: string, messageId: string, attachmentId: string): Promise<Uint8Array> {
  const res = await fetch(
    `${GRAPH}${box}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (!res.ok) throw new Error(`Graph GET attachment failed: ${res.status} ${await res.text()}`);
  return new Uint8Array(await res.arrayBuffer());
}

// Sends an HTML email from the connected mailbox (saved to its Sent Items).
async function sendMail(token: string, box: string, to: string, subject: string, html: string): Promise<void> {
  await graph<void>(token, `${box}/sendMail`, {
    method: "POST",
    body: JSON.stringify({
      message: { subject, body: { contentType: "HTML", content: html }, toRecipients: [{ emailAddress: { address: to } }] },
      saveToSentItems: true,
    }),
  });
}

async function forwardMessage(token: string, box: string, messageId: string, to: string, comment: string): Promise<void> {
  await graph<void>(token, `${box}/messages/${encodeURIComponent(messageId)}/forward`, {
    method: "POST",
    body: JSON.stringify({ comment, toRecipients: [{ emailAddress: { address: to } }] }),
  });
}
