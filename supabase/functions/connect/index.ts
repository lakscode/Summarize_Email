// Connect a mailbox.
//   POST {action: "providers"}                            -> which options the admin has configured
//   POST {provider: "microsoft" | "gmail"}                -> { url } to start OAuth consent
//   GET  ?code=...&state=...                               -> OAuth callback, stores the refresh token
//   POST {provider: "microsoft_app", email}               -> shared mailbox via app-only Graph permissions
//   POST {provider: "imap", email, imap_host, imap_port, smtp_host, smtp_port, username, password}
import type { SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import { GOOGLE_SCOPES } from "../_shared/providers/gmail.ts";
import { MicrosoftProvider, MS_SCOPES, msTokenUrl } from "../_shared/providers/microsoft.ts";
import { ImapProvider, type ImapCredentials } from "../_shared/providers/imap.ts";
import { writeSecret } from "../_shared/providers/index.ts";
import type { ProviderKind } from "../_shared/types.ts";
import { adminClient, b64, callerUserId, corsHeaders, env, json } from "../_shared/util.ts";

const redirectUri = () => Deno.env.get("OAUTH_REDIRECT_URI") ?? `${env("SUPABASE_URL")}/functions/v1/connect`;
const appPage = (name: string) =>
  `${env("APP_URL", "http://localhost:5500").replace(/\/(index\.html)?$/, "")}/${name}`;
const has = (name: string) => !!Deno.env.get(name);

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  const url = new URL(req.url);

  try {
    if (req.method === "GET" && (url.searchParams.has("code") || url.searchParams.has("error"))) {
      return await oauthCallback(url);
    }
    if (req.method !== "POST") return json({ error: "Not found" }, 404);

    const userId = await callerUserId(req);
    if (!userId) return json({ error: "Unauthorized" }, 401);
    const body = await req.json();

    if (body.action === "providers") {
      // Which connection options the administrator has configured (no secrets returned).
      const ms = has("MS_CLIENT_ID") && has("MS_CLIENT_SECRET");
      const tenant = Deno.env.get("MS_TENANT_ID") ?? "common";
      return json({
        microsoft: ms,
        microsoft_app: ms && !["common", "organizations", "consumers"].includes(tenant),
        gmail: has("GOOGLE_CLIENT_ID") && has("GOOGLE_CLIENT_SECRET"),
        imap: true,
        redirect_uri: redirectUri(),
      });
    }

    switch (body.provider) {
      case "microsoft":
        return json({
          url: `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0/authorize?` +
            new URLSearchParams({
              client_id: env("MS_CLIENT_ID"),
              response_type: "code",
              redirect_uri: redirectUri(),
              response_mode: "query",
              scope: MS_SCOPES,
              prompt: "select_account",
              state: await signState({ uid: userId, provider: "microsoft" }),
            }),
        });
      case "gmail":
        return json({
          url: "https://accounts.google.com/o/oauth2/v2/auth?" +
            new URLSearchParams({
              client_id: env("GOOGLE_CLIENT_ID"),
              response_type: "code",
              redirect_uri: redirectUri(),
              scope: GOOGLE_SCOPES,
              access_type: "offline",
              prompt: "consent",
              state: await signState({ uid: userId, provider: "gmail" }),
            }),
        });
      case "microsoft_app": {
        const email = String(body.email ?? "").trim().toLowerCase();
        if (!email.includes("@")) return json({ error: "Mailbox email address is required" }, 400);
        // Prove the app registration can read this mailbox before saving it.
        await new MicrosoftProvider("app", email, null).listMessages(new Date(Date.now() - 3600_000));
        const id = await saveMailbox(adminClient(), userId, "microsoft_app", email, null, null);
        return json({ mailbox_id: id, email });
      }
      case "imap": {
        const email = String(body.email ?? "").trim().toLowerCase();
        const creds: ImapCredentials = {
          imap_host: body.imap_host,
          imap_port: Number(body.imap_port ?? 993),
          smtp_host: body.smtp_host,
          smtp_port: Number(body.smtp_port ?? 465),
          username: body.username || email,
          password: body.password,
          inbox: body.inbox || "INBOX",
        };
        if (!email.includes("@") || !creds.imap_host || !creds.smtp_host || !creds.password) {
          return json({ error: "email, imap_host, smtp_host and password are required" }, 400);
        }
        const probe = new ImapProvider(email, creds);
        try {
          await probe.listMessages(new Date(Date.now() - 3600_000));
        } finally {
          await probe.close();
        }
        const id = await saveMailbox(adminClient(), userId, "imap", email, null, JSON.stringify(creds));
        return json({ mailbox_id: id, email });
      }
      default:
        return json({ error: "provider must be microsoft, microsoft_app, gmail or imap" }, 400);
    }
  } catch (err) {
    console.error(err);
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

async function oauthCallback(url: URL): Promise<Response> {
  const state = await verifyState(url.searchParams.get("state") ?? "");
  const back = (params: Record<string, string>) =>
    Response.redirect(`${appPage("connect.html")}?${new URLSearchParams({ ...(state ? { provider: state.provider } : {}), ...params })}`, 302);

  if (url.searchParams.get("error")) {
    return back({ connect_error: url.searchParams.get("error_description") ?? url.searchParams.get("error")! });
  }
  if (!state) return back({ connect_error: "Invalid or expired sign-in state - please try again." });
  const code = url.searchParams.get("code")!;

  let email: string;
  let name: string | null = null;
  let refreshToken: string | undefined;

  if (state.provider === "microsoft") {
    const res = await fetch(msTokenUrl(), {
      method: "POST",
      body: new URLSearchParams({
        client_id: env("MS_CLIENT_ID"),
        client_secret: env("MS_CLIENT_SECRET"),
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri(),
        scope: MS_SCOPES,
      }),
    });
    if (!res.ok) return back({ connect_error: `Microsoft sign-in failed: ${await res.text()}` });
    const tok = await res.json();
    refreshToken = tok.refresh_token;
    const me = await (await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName", {
      headers: { Authorization: `Bearer ${tok.access_token}` },
    })).json();
    email = (me.mail ?? me.userPrincipalName ?? "").toLowerCase();
    name = me.displayName ?? null;
  } else {
    const res = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      body: new URLSearchParams({
        client_id: env("GOOGLE_CLIENT_ID"),
        client_secret: env("GOOGLE_CLIENT_SECRET"),
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri(),
      }),
    });
    if (!res.ok) return back({ connect_error: `Google sign-in failed: ${await res.text()}` });
    const tok = await res.json();
    refreshToken = tok.refresh_token;
    const profile = await (await fetch("https://gmail.googleapis.com/gmail/v1/users/me/profile", {
      headers: { Authorization: `Bearer ${tok.access_token}` },
    })).json();
    email = (profile.emailAddress ?? "").toLowerCase();
  }

  if (!refreshToken) return back({ connect_error: "No refresh token was returned - remove the app's access and connect again." });
  if (!email) return back({ connect_error: "Could not read the mailbox address." });

  await saveMailbox(adminClient(), state.uid, state.provider, email, name, refreshToken);
  return back({ connected: email });
}

async function saveMailbox(
  db: SupabaseClient,
  userId: string,
  provider: ProviderKind,
  email: string,
  name: string | null,
  secret: string | null,
): Promise<string> {
  const { data: existing } = await db.from("mailboxes").select("id, credentials_secret_id")
    .eq("owner_user_id", userId).eq("email_address", email).maybeSingle();
  let id = existing?.id as string | undefined;
  if (!id) {
    const { data, error } = await db.from("mailboxes")
      .insert({ owner_user_id: userId, provider, email_address: email, display_name: name })
      .select("id").single();
    if (error) throw new Error(error.message);
    id = data.id as string;
  }
  const update: Record<string, unknown> = { provider, active: true };
  if (name) update.display_name = name;
  if (secret) update.credentials_secret_id = await writeSecret(db, existing?.credentials_secret_id ?? null, secret, `mailbox:${id}`);
  const { error } = await db.from("mailboxes").update(update).eq("id", id);
  if (error) throw new Error(error.message);
  return id;
}

// --- signed OAuth state (HMAC-SHA256, 15 minute expiry) ----------------------------------

interface State {
  uid: string;
  provider: "microsoft" | "gmail";
  exp?: number;
}

async function hmacKey() {
  return await crypto.subtle.importKey("raw", new TextEncoder().encode(env("OAUTH_STATE_SECRET")),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
}

async function signState(s: State): Promise<string> {
  const payload = b64.encodeUrl(new TextEncoder().encode(JSON.stringify({ ...s, exp: Date.now() + 15 * 60_000 })));
  const sig = new Uint8Array(await crypto.subtle.sign("HMAC", await hmacKey(), new TextEncoder().encode(payload)));
  return `${payload}.${b64.encodeUrl(sig)}`;
}

async function verifyState(token: string): Promise<State | null> {
  const [payload, sig] = token.split(".");
  if (!payload || !sig) return null;
  const ok = await crypto.subtle.verify("HMAC", await hmacKey(), b64.decodeUrl(sig), new TextEncoder().encode(payload));
  if (!ok) return null;
  const s = JSON.parse(new TextDecoder().decode(b64.decodeUrl(payload))) as State;
  return s.exp && s.exp > Date.now() ? s : null;
}
