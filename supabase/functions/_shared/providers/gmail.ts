// Gmail / Google Workspace via the Gmail REST API (OAuth refresh token stored in Vault).
import type { AttachmentMeta, MailMessage, MailProvider, MailSummary } from "../types.ts";
import { buildMime } from "../mime.ts";
import { b64, env, htmlToText } from "../util.ts";

const API = "https://gmail.googleapis.com/gmail/v1/users/me";
export const GOOGLE_SCOPES = "https://www.googleapis.com/auth/gmail.modify https://www.googleapis.com/auth/gmail.send openid email";

interface GmailPart {
  partId?: string;
  mimeType: string;
  filename?: string;
  headers?: { name: string; value: string }[];
  body?: { attachmentId?: string; data?: string; size?: number };
  parts?: GmailPart[];
}

export class GmailProvider implements MailProvider {
  private accessToken: string | null = null;
  private labelIds = new Map<string, string>();
  private me: string;

  constructor(mailbox: string, private refreshToken: string) {
    this.me = mailbox;
  }

  private async token(): Promise<string> {
    if (this.accessToken) return this.accessToken;
    const res = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      body: new URLSearchParams({
        client_id: env("GOOGLE_CLIENT_ID"),
        client_secret: env("GOOGLE_CLIENT_SECRET"),
        refresh_token: this.refreshToken,
        grant_type: "refresh_token",
      }),
    });
    if (!res.ok) throw new Error(`Google token request failed: ${res.status} ${await res.text()}`);
    this.accessToken = (await res.json()).access_token;
    return this.accessToken!;
  }

  private async api<T>(path: string, init: RequestInit = {}): Promise<T> {
    for (let attempt = 0; ; attempt++) {
      const res = await fetch(`${API}${path}`, {
        ...init,
        headers: { Authorization: `Bearer ${await this.token()}`, "Content-Type": "application/json" },
      });
      if ((res.status === 429 || res.status >= 500) && attempt < 3) {
        await new Promise((r) => setTimeout(r, 1000 * 2 ** attempt));
        continue;
      }
      if (!res.ok) throw new Error(`Gmail ${init.method ?? "GET"} ${path} failed: ${res.status} ${await res.text()}`);
      return (await res.json()) as T;
    }
  }

  async listMessages(since: Date): Promise<MailSummary[]> {
    const ids: string[] = [];
    const q = encodeURIComponent(`in:inbox -from:me after:${Math.floor(since.getTime() / 1000)}`);
    let pageToken = "";
    do {
      const page: { messages?: { id: string }[]; nextPageToken?: string } = await this.api(
        `/messages?q=${q}&maxResults=100${pageToken ? `&pageToken=${pageToken}` : ""}`,
      );
      ids.push(...(page.messages ?? []).map((m) => m.id));
      pageToken = page.nextPageToken ?? "";
    } while (pageToken);

    const labelNames = await this.labelNameMap();
    const out: MailSummary[] = [];
    for (const id of ids) {
      const m: any = await this.api(`/messages/${id}?format=minimal`);
      out.push({
        id,
        threadId: m.threadId ?? null,
        receivedAt: new Date(Number(m.internalDate)),
        labels: (m.labelIds ?? []).map((l: string) => labelNames.get(l) ?? l),
      });
    }
    return out;
  }

  private async labelNameMap(): Promise<Map<string, string>> {
    const res: { labels: { id: string; name: string }[] } = await this.api(`/labels`);
    const map = new Map<string, string>();
    for (const l of res.labels ?? []) {
      map.set(l.id, l.name);
      this.labelIds.set(l.name, l.id);
    }
    return map;
  }

  async getMessage(id: string): Promise<MailMessage> {
    const m: any = await this.api(`/messages/${id}?format=full`);
    const headers: Record<string, string> = {};
    for (const h of m.payload?.headers ?? []) {
      const k = h.name.toLowerCase();
      headers[k] = headers[k] ? `${headers[k]}\n${h.value}` : h.value;
    }
    let text = "";
    let html = "";
    const attachments: AttachmentMeta[] = [];
    const walk = (p: GmailPart) => {
      if (p.filename && (p.body?.attachmentId || p.body?.data)) {
        const disp = p.headers?.find((h) => h.name.toLowerCase() === "content-disposition")?.value ?? "";
        if (!/^inline/i.test(disp) || /pdf|sheet|excel|csv/i.test(p.mimeType)) {
          attachments.push({
            id: p.body.attachmentId ?? `inline-${p.partId}`,
            name: p.filename,
            contentType: p.mimeType,
            size: p.body.size ?? 0,
            content: p.body.data ? b64.decodeUrl(p.body.data) : undefined,
          });
        }
      } else if (p.mimeType === "text/plain" && p.body?.data && !text) {
        text = new TextDecoder().decode(b64.decodeUrl(p.body.data));
      } else if (p.mimeType === "text/html" && p.body?.data && !html) {
        html = new TextDecoder().decode(b64.decodeUrl(p.body.data));
      }
      for (const c of p.parts ?? []) walk(c);
    };
    walk(m.payload);
    const labelNames = this.labelIds.size ? new Map([...this.labelIds].map(([n, i]) => [i, n])) : await this.labelNameMap();
    const from = parseAddress(headers["from"] ?? "");
    return {
      id,
      threadId: m.threadId ?? null,
      from,
      replyTo: headers["reply-to"] ? headers["reply-to"].split(",").map((a) => parseAddress(a).address) : [],
      to: (headers["to"] ?? "").split(",").map((a) => parseAddress(a).address).filter(Boolean),
      subject: headers["subject"] ?? "",
      receivedAt: new Date(Number(m.internalDate)),
      bodyText: text || htmlToText(html),
      headers,
      labels: (m.labelIds ?? []).map((l: string) => labelNames.get(l) ?? l),
      attachments,
    };
  }

  async getAttachment(messageId: string, att: AttachmentMeta): Promise<Uint8Array> {
    if (att.content) return att.content;
    const a: { data: string } = await this.api(`/messages/${messageId}/attachments/${att.id}`);
    return b64.decodeUrl(a.data);
  }

  async forward(message: MailMessage, to: string, comment: string): Promise<void> {
    const attachments = [];
    for (const a of message.attachments) {
      attachments.push({ name: a.name, contentType: a.contentType, content: await this.getAttachment(message.id, a) });
    }
    const quoted =
      `${comment}\n\n---------- Forwarded message ---------\nFrom: ${message.from.name ?? ""} <${message.from.address}>\n` +
      `Date: ${message.receivedAt.toUTCString()}\nSubject: ${message.subject}\n\n${message.bodyText}`;
    await this.sendRaw(buildMime({ from: this.me, to, subject: `Fwd: ${message.subject}`, text: quoted, attachments }));
  }

  async sendHtml(to: string, subject: string, html: string): Promise<void> {
    await this.sendRaw(buildMime({ from: this.me, to, subject, html, text: htmlToText(html) }));
  }

  private async sendRaw(mime: string): Promise<void> {
    await this.api(`/messages/send`, {
      method: "POST",
      body: JSON.stringify({ raw: b64.encodeUrl(new TextEncoder().encode(mime)) }),
    });
  }

  async addLabel(messageId: string, label: string): Promise<void> {
    if (!this.labelIds.size) await this.labelNameMap();
    let labelId = this.labelIds.get(label);
    if (!labelId) {
      const created: { id: string } = await this.api(`/labels`, {
        method: "POST",
        body: JSON.stringify({ name: label, labelListVisibility: "labelShow", messageListVisibility: "show" }),
      });
      labelId = created.id;
      this.labelIds.set(label, labelId);
    }
    await this.api(`/messages/${messageId}/modify`, { method: "POST", body: JSON.stringify({ addLabelIds: [labelId] }) });
  }
}

export function parseAddress(raw: string): { address: string; name: string | null } {
  const m = raw.match(/^\s*"?([^"<]*)"?\s*<([^>]+)>/);
  if (m) return { name: m[1].trim() || null, address: m[2].trim().toLowerCase() };
  return { name: null, address: raw.trim().toLowerCase() };
}
