import type { SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import type { Mailbox, MailProvider } from "../types.ts";
import { GmailProvider } from "./gmail.ts";
import { ImapProvider, type ImapCredentials } from "./imap.ts";
import { MicrosoftProvider } from "./microsoft.ts";

export async function readSecret(db: SupabaseClient, secretId: string | null): Promise<string | null> {
  if (!secretId) return null;
  const { data, error } = await db.rpc("mailbox_secret_get", { p_secret_id: secretId });
  if (error) throw new Error(`Could not read mailbox credentials: ${error.message}`);
  return data as string | null;
}

export async function writeSecret(db: SupabaseClient, secretId: string | null, value: string, name: string): Promise<string> {
  const { data, error } = await db.rpc("mailbox_secret_put", { p_secret_id: secretId, p_secret: value, p_name: name });
  if (error) throw new Error(`Could not store mailbox credentials: ${error.message}`);
  return data as string;
}

export async function providerFor(db: SupabaseClient, mailbox: Mailbox): Promise<MailProvider> {
  const secret = await readSecret(db, mailbox.credentials_secret_id);
  switch (mailbox.provider) {
    case "microsoft":
      return new MicrosoftProvider("delegated", mailbox.email_address, secret, async (rotated) => {
        await writeSecret(db, mailbox.credentials_secret_id, rotated, `mailbox:${mailbox.id}`);
      });
    case "microsoft_app":
      return new MicrosoftProvider("app", mailbox.email_address, null);
    case "gmail":
      if (!secret) throw new Error("Gmail mailbox has no refresh token; reconnect it.");
      return new GmailProvider(mailbox.email_address, secret);
    case "imap":
      if (!secret) throw new Error("IMAP mailbox has no credentials; reconnect it.");
      return new ImapProvider(mailbox.email_address, JSON.parse(secret) as ImapCredentials);
  }
}
