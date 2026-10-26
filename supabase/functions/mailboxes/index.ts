// Azure / Microsoft 365 mailboxes for the signed-in user.
//   GET                         -> { mailboxes: [...] }
//   POST {action: "connect"}    -> { url } to start Microsoft sign-in (delegated OAuth)
//   DELETE ?id=<mailbox id>     -> disconnect (Vault secret removed by trigger)
//   GET ?code=...&state=...     -> OAuth callback from Microsoft; stores the refresh token, redirects to the dashboard
// verify_jwt is off for this function (the Microsoft callback carries no Supabase JWT); every other
// route checks the caller itself.
import { decodeBase64Url, encodeBase64Url } from "jsr:@std/encoding@1/base64url";
import { corsHeaders, json } from "../_shared/cors.ts";
import { adminClient, callerUser, env } from "../_shared/supabase.ts";

const MS_SCOPES = "offline_access User.Read Mail.Read";
const msBase = () => `https://login.microsoftonline.com/${env("MS_TENANT_ID", "common")}/oauth2/v2.0`;
const redirectUri = () => `${env("SUPABASE_URL")}/functions/v1/mailboxes`;
const dashboardUrl = () => `${env("APP_URL").replace(/\/+$/, "")}/dashboard`;

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  const url = new URL(req.url);

  try {
    if (req.method === "GET" && (url.searchParams.has("code") || url.searchParams.has("error"))) {
      return await oauthCallback(url);
    }

    const caller = await callerUser(req);
    if (!caller) return json({ error: "Not signed in" }, 401);
    const { user, sb } = caller;

    if (req.method === "GET") {
      const { data, error } = await sb
        .from("mailboxes")
        .select("id, provider, email_address, display_name, active, created_at")
        .order("created_at");
      if (error) return json({ error: error.message }, 500);
      return json({ mailboxes: data });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action !== "connect") return json({ error: "Unknown action" }, 400);
      if (!Deno.env.get("MS_CLIENT_ID") || !Deno.env.get("MS_CLIENT_SECRET")) {
        return json({ error: "Azure isn't set up on the server yet (MS_CLIENT_ID / MS_CLIENT_SECRET are missing)." }, 400);
      }
      return json({
        url: `${msBase()}/authorize?` + new URLSearchParams({
          client_id: env("MS_CLIENT_ID"),
          response_type: "code",
          redirect_uri: redirectUri(),
          response_mode: "query",
          scope: MS_SCOPES,
          prompt: "select_account",
          state: await signState(user.id),
        }),
      });
    }

    if (req.method === "DELETE") {
      const id = url.searchParams.get("id");
      if (!id) return json({ error: "id is required" }, 400);
      const { error, count } = await sb.from("mailboxes").delete({ count: "exact" }).eq("id", id);
      if (error) return json({ error: error.message }, 500);
      if (!count) return json({ error: "Mailbox not found" }, 404);
      return json({ deleted: id });
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    console.error(err);
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

async function oauthCallback(url: URL): Promise<Response> {
  const back = (params: Record<string, string>) =>
    Response.redirect(`${dashboardUrl()}?${new URLSearchParams(params)}`, 302);

  if (url.searchParams.get("error")) {
    return back({ connect_error: url.searchParams.get("error_description") ?? url.searchParams.get("error")! });
  }
  const userId = await verifyState(url.searchParams.get("state") ?? "");
  if (!userId) return back({ connect_error: "Sign-in expired or was tampered with. Please try again." });

  const tokenRes = await fetch(`${msBase()}/token`, {
    method: "POST",
    body: new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      client_secret: env("MS_CLIENT_SECRET"),
      grant_type: "authorization_code",
      code: url.searchParams.get("code")!,
      redirect_uri: redirectUri(),
      scope: MS_SCOPES,
    }),
  });
  if (!tokenRes.ok) {
    console.error("Microsoft token exchange failed", await tokenRes.text());
    return back({ connect_error: "Microsoft sign-in failed. Please try again." });
  }
  const tok = await tokenRes.json();
  if (!tok.refresh_token) return back({ connect_error: "Microsoft didn't return a refresh token. Remove the app's access and connect again." });

  const me = await (await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName", {
    headers: { Authorization: `Bearer ${tok.access_token}` },
  })).json();
  const email = String(me.mail ?? me.userPrincipalName ?? "").toLowerCase();
  if (!email) return back({ connect_error: "Could not read the mailbox address." });

  const db = adminClient();
  const { data: existing } = await db.from("mailboxes").select("id, credentials_secret_id")
    .eq("owner_user_id", userId).eq("email_address", email).maybeSingle();

  let id = existing?.id as string | undefined;
  if (!id) {
    const { data, error } = await db.from("mailboxes")
      .insert({ owner_user_id: userId, provider: "microsoft", email_address: email, display_name: me.displayName ?? null })
      .select("id").single();
    if (error) throw new Error(error.message);
    id = data.id as string;
  }

  const { data: secretId, error: secretError } = await db.rpc("mailbox_secret_put", {
    p_secret_id: existing?.credentials_secret_id ?? null,
    p_secret: tok.refresh_token,
    p_name: `mailbox:${id}`,
  });
  if (secretError) throw new Error(`Could not store mailbox credentials: ${secretError.message}`);

  const { error } = await db.from("mailboxes").update({
    credentials_secret_id: secretId,
    display_name: me.displayName ?? null,
    active: true,
    updated_at: new Date().toISOString(),
  }).eq("id", id);
  if (error) throw new Error(error.message);

  return back({ connected: email });
}

// --- signed OAuth state (HMAC-SHA256, 15 minute expiry) ----------------------------------

const enc = new TextEncoder();

async function hmacKey() {
  return await crypto.subtle.importKey("raw", enc.encode(env("OAUTH_STATE_SECRET")),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
}

async function signState(uid: string): Promise<string> {
  const payload = encodeBase64Url(enc.encode(JSON.stringify({ uid, exp: Date.now() + 15 * 60_000 })));
  const sig = new Uint8Array(await crypto.subtle.sign("HMAC", await hmacKey(), enc.encode(payload)));
  return `${payload}.${encodeBase64Url(sig)}`;
}

async function verifyState(token: string): Promise<string | null> {
  const [payload, sig] = token.split(".");
  if (!payload || !sig) return null;
  const ok = await crypto.subtle.verify("HMAC", await hmacKey(), decodeBase64Url(sig), enc.encode(payload));
  if (!ok) return null;
  const s = JSON.parse(new TextDecoder().decode(decodeBase64Url(payload)));
  return s.exp > Date.now() && typeof s.uid === "string" ? s.uid : null;
}
