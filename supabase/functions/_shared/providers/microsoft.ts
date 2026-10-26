// Microsoft 365 / Azure (Outlook) via Microsoft Graph.
//  - "microsoft":     delegated OAuth, refresh token stored in Vault, acts as /me
//  - "microsoft_app": app-only client credentials (Mail.ReadWrite + Mail.Send application
//                     permissions), acts on /users/{mailbox} - suits shared AP mailboxes.
import type { AttachmentMeta, MailMessage, MailProvider, MailSummary } from "../types.ts";
import { b64, env, htmlToText } from "../util.ts";

const GRAPH = "https://graph.microsoft.com/v1.0";
export const MS_SCOPES = "offline_access User.Read Mail.ReadWrite Mail.Send";

export function msTokenUrl(): string {
  return `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0/token`;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
}

export class MicrosoftProvider implements MailProvider {
  private accessToken: string | null = null;
  private base: string;

  constructor(
    private mode: "delegated" | "app",
    mailbox: string,
    private refreshToken: string | null,
    /** Called when Microsoft rotates the refresh token so it can be persisted. */
    private onRefreshToken?: (token: string) => Promise<void>,
  ) {
    this.base = mode === "app" ? `${GRAPH}/users/${encodeURIComponent(mailbox)}` : `${GRAPH}/me`;
  }

  private async token(): Promise<string> {
    if (this.accessToken) return this.accessToken;
    const form = new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
    });
    if (this.mode === "app") {
      form.set("grant_type", "client_credentials");
      form.set("scope", "https://graph.microsoft.com/.default");
    } else {
      if (!this.refreshToken) throw new Error("Mailbox has no Microsoft refresh token; reconnect it.");
      form.set("grant_type", "refresh_token");
      form.set("refresh_token", this.refreshToken);
      form.set("scope", MS_SCOPES);
    }
    const res = await fetch(msTokenUrl(), { method: "POST", body: form });
    if (!res.ok) throw new Error(`Microsoft token request failed: ${res.status} ${await res.text()}`);
    const tok = (await res.json()) as TokenResponse;
    this.accessToken = tok.access_token;
    if (tok.refresh_token && tok.refresh_token !== this.refreshToken && this.onRefreshToken) {
      this.refreshToken = tok.refresh_token;
      await this.onRefreshToken(tok.refresh_token);
    }
    return this.accessToken;
  }

  private async graph<T>(path: string, init: RequestInit = {}, extraHeaders: Record<string, string> = {}): Promise<T> {
    const url = path.startsWith("http") ? path : `${this.base}${path}`;
    for (let attempt = 0; ; attempt++) {
      const res = await fetch(url, {
        ...init,
        headers: {
          Authorization: `Bearer ${await this.token()}`,
          "Content-Type": "application/json",
          ...extraHeaders,
        },
      });
      if (res.status === 429 && attempt < 3) {
        await new Promise((r) => setTimeout(r, Number(res.headers.get("Retry-After") ?? 2) * 1000));
        continue;
      }
      if (!res.ok) throw new Error(`Graph ${init.method ?? "GET"} ${path} failed: ${res.status} ${await res.text()}`);
      if (res.status === 202 || res.status === 204) return undefined as T;
      return (await res.json()) as T;
    }
  }

  async listMessages(since: Date): Promise<MailSummary[]> {
    const out: MailSummary[] = [];
    let next: string | null =
      `/mailFolders/inbox/messages?$filter=receivedDateTime ge ${since.toISOString()}` +
      `&$select=id,conversationId,receivedDateTime,categories,from&$orderby=receivedDateTime desc&$top=50`;
    while (next) {
      const page: { value: any[]; "@odata.nextLink"?: string } = await this.graph(next);
      for (const m of page.value) {
        out.push({
          id: m.id,
          threadId: m.conversationId ?? null,
          receivedAt: new Date(m.receivedDateTime),
          labels: m.categories ?? [],
        });
      }
      next = page["@odata.nextLink"] ?? null;
    }
    return out;
  }

  async getMessage(id: string): Promise<MailMessage> {
    const m: any = await this.graph(
      `/messages/${id}?$select=id,conversationId,subject,from,replyTo,toRecipients,receivedDateTime,body,categories,internetMessageHeaders`,
      {},
      { Prefer: 'outlook.body-content-type="text"' },
    );
    const atts: { value: any[] } = await this.graph(`/messages/${id}/attachments`);
    const headers: Record<string, string> = {};
    for (const h of m.internetMessageHeaders ?? []) headers[h.name.toLowerCase()] = h.value;
    return {
      id: m.id,
      threadId: m.conversationId ?? null,
      from: { address: m.from?.emailAddress?.address ?? "", name: m.from?.emailAddress?.name ?? null },
      replyTo: (m.replyTo ?? []).map((r: any) => r.emailAddress?.address).filter(Boolean),
      to: (m.toRecipients ?? []).map((r: any) => r.emailAddress?.address).filter(Boolean),
      subject: m.subject ?? "",
      receivedAt: new Date(m.receivedDateTime),
      bodyText: m.body?.contentType === "html" ? htmlToText(m.body.content ?? "") : (m.body?.content ?? ""),
      headers,
      labels: m.categories ?? [],
      attachments: atts.value
        .filter((a) => a["@odata.type"] === "#microsoft.graph.fileAttachment" && !a.isInline)
        .map((a) => ({
          id: a.id,
          name: a.name,
          contentType: a.contentType ?? "application/octet-stream",
          size: a.size ?? 0,
          content: a.contentBytes ? b64.decode(a.contentBytes) : undefined,
        })),
    };
  }

  async getAttachment(messageId: string, att: AttachmentMeta): Promise<Uint8Array> {
    if (att.content) return att.content;
    const a: any = await this.graph(`/messages/${messageId}/attachments/${att.id}`);
    return b64.decode(a.contentBytes);
  }

  async forward(message: MailMessage, to: string, comment: string): Promise<void> {
    await this.graph(`/messages/${message.id}/forward`, {
      method: "POST",
      body: JSON.stringify({ comment, toRecipients: [{ emailAddress: { address: to } }] }),
    });
  }

  async sendHtml(to: string, subject: string, html: string): Promise<void> {
    await this.graph(`/sendMail`, {
      method: "POST",
      body: JSON.stringify({
        message: { subject, body: { contentType: "HTML", content: html }, toRecipients: [{ emailAddress: { address: to } }] },
        saveToSentItems: true,
      }),
    });
  }

  async addLabel(messageId: string, label: string): Promise<void> {
    const m: any = await this.graph(`/messages/${messageId}?$select=categories`);
    const categories: string[] = m.categories ?? [];
    if (categories.includes(label)) return;
    await this.graph(`/messages/${messageId}`, {
      method: "PATCH",
      body: JSON.stringify({ categories: [...categories, label] }),
    });
  }
}
