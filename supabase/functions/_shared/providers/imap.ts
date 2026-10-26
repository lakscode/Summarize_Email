// Any other mailbox over IMAP (read) + SMTP (send).
// Note: Supabase Edge Functions block outbound ports 25 and 587 - use SMTP over TLS on 465.
import { Buffer } from "node:buffer";
import { ImapFlow } from "npm:imapflow@2.9.1";
import nodemailer from "npm:nodemailer@10.2.1";
import { simpleParser } from "npm:mailparser@3.12.4";
import type { AttachmentMeta, MailMessage, MailProvider, MailSummary } from "../types.ts";
import { htmlToText } from "../util.ts";

export interface ImapCredentials {
  imap_host: string;
  imap_port: number;
  smtp_host: string;
  smtp_port: number;
  username: string;
  password: string;
  inbox?: string;
}

export class ImapProvider implements MailProvider {
  private client: ImapFlow | null = null;
  private cache = new Map<string, MailMessage>();
  private inbox: string;

  constructor(private mailbox: string, private creds: ImapCredentials) {
    this.inbox = creds.inbox ?? "INBOX";
  }

  private async imap(): Promise<ImapFlow> {
    if (this.client) return this.client;
    this.client = new ImapFlow({
      host: this.creds.imap_host,
      port: this.creds.imap_port ?? 993,
      secure: true,
      auth: { user: this.creds.username, pass: this.creds.password },
      logger: false,
    });
    await this.client.connect();
    return this.client;
  }

  async listMessages(since: Date): Promise<MailSummary[]> {
    const c = await this.imap();
    const lock = await c.getMailboxLock(this.inbox);
    try {
      const out: MailSummary[] = [];
      const uids = (await c.search({ since }, { uid: true })) || [];
      if (!uids.length) return out;
      for await (const m of c.fetch(uids, { uid: true, envelope: true, flags: true, internalDate: true, threadId: true }, { uid: true })) {
        out.push({
          id: String(m.uid),
          threadId: m.threadId ?? m.envelope?.inReplyTo ?? m.envelope?.messageId ?? null,
          receivedAt: new Date(m.internalDate ?? m.envelope?.date ?? Date.now()),
          labels: [...(m.flags ?? [])],
        });
      }
      return out;
    } finally {
      lock.release();
    }
  }

  async getMessage(id: string): Promise<MailMessage> {
    const cached = this.cache.get(id);
    if (cached) return cached;
    const c = await this.imap();
    const lock = await c.getMailboxLock(this.inbox);
    try {
      const m = await c.fetchOne(id, { source: true, flags: true, internalDate: true }, { uid: true });
      if (!m || !m.source) throw new Error(`IMAP message ${id} not found`);
      const parsed = await simpleParser(m.source);
      const headers: Record<string, string> = {};
      for (const line of parsed.headerLines ?? []) {
        const k = line.key.toLowerCase();
        const v = line.line.slice(line.line.indexOf(":") + 1).trim();
        headers[k] = headers[k] ? `${headers[k]}\n${v}` : v;
      }
      const from = parsed.from?.value?.[0];
      const replyTo = parsed.replyTo?.value?.map((a: any) => a.address).filter(Boolean) ?? [];
      const toList = Array.isArray(parsed.to) ? parsed.to.flatMap((t: any) => t.value) : (parsed.to?.value ?? []);
      const msg: MailMessage = {
        id,
        threadId: (parsed.references?.[0] as string | undefined) ?? parsed.inReplyTo ?? parsed.messageId ?? null,
        from: { address: (from?.address ?? "").toLowerCase(), name: from?.name || null },
        replyTo,
        to: toList.map((a: any) => a.address).filter(Boolean),
        subject: parsed.subject ?? "",
        receivedAt: new Date(m.internalDate ?? parsed.date ?? Date.now()),
        bodyText: parsed.text || htmlToText(typeof parsed.html === "string" ? parsed.html : ""),
        headers,
        labels: [...(m.flags ?? [])],
        attachments: (parsed.attachments ?? [])
          .filter((a: any) => a.contentDisposition !== "inline" || /pdf|sheet|excel|csv/i.test(a.contentType))
          .map((a: any, i: number): AttachmentMeta => ({
            id: String(i),
            name: a.filename ?? `attachment-${i}`,
            contentType: a.contentType,
            size: a.size,
            content: new Uint8Array(a.content),
          })),
      };
      this.cache.set(id, msg);
      return msg;
    } finally {
      lock.release();
    }
  }

  getAttachment(_messageId: string, att: AttachmentMeta): Promise<Uint8Array> {
    if (!att.content) throw new Error(`Attachment ${att.name} has no content`);
    return Promise.resolve(att.content);
  }

  private transport() {
    return nodemailer.createTransport({
      host: this.creds.smtp_host,
      port: this.creds.smtp_port ?? 465,
      secure: (this.creds.smtp_port ?? 465) === 465,
      auth: { user: this.creds.username, pass: this.creds.password },
    });
  }

  async forward(message: MailMessage, to: string, comment: string): Promise<void> {
    await this.transport().sendMail({
      from: this.mailbox,
      to,
      subject: `Fwd: ${message.subject}`,
      text:
        `${comment}\n\n---------- Forwarded message ---------\nFrom: ${message.from.name ?? ""} <${message.from.address}>\n` +
        `Date: ${message.receivedAt.toUTCString()}\nSubject: ${message.subject}\n\n${message.bodyText}`,
      attachments: message.attachments.map((a) => ({ filename: a.name, contentType: a.contentType, content: a.content ? Buffer.from(a.content) : undefined })),
    });
  }

  async sendHtml(to: string, subject: string, html: string): Promise<void> {
    await this.transport().sendMail({ from: this.mailbox, to, subject, html, text: htmlToText(html) });
  }

  async addLabel(messageId: string, label: string): Promise<void> {
    const c = await this.imap();
    const lock = await c.getMailboxLock(this.inbox);
    try {
      // IMAP keywords cannot contain spaces or some specials; "AI-Processed" is fine.
      await c.messageFlagsAdd(messageId, [label.replace(/[^A-Za-z0-9_$.-]/g, "_")], { uid: true });
    } finally {
      lock.release();
    }
  }

  async close(): Promise<void> {
    await this.client?.logout().catch(() => {});
    this.client = null;
  }
}
