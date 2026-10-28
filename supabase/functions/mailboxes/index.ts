// Connected mailboxes (Azure / Microsoft 365 or Gmail) for the signed-in user.
//   GET                                   -> { mailboxes: [...] } (expired access is cleaned up first)
//   POST {action: "connect", provider: "microsoft" | "google", access_days, popup, shared_mailbox?}
//                                         -> { url } to start Microsoft or Google sign-in (OAuth). With
//                                            shared_mailbox (Microsoft only), the user signs in with their own
//                                            account and the app reads that shared mailbox through it.
//   POST {action: "connect_app", email_address, tenant_id, client_id, client_secret, access_days}
//                                         -> { mailbox } connect a Microsoft mailbox (usually shared) with API
//                                            keys: the customer's own Azure app registration with application
//                                            permissions Mail.Read and Mail.Send. Checked before saving; the keys
//                                            are kept in Vault and never returned.
//   PATCH {id, smart_ap_address}          -> set where invoices are forwarded ("" clears it)
//   PATCH {id, schedule: {enabled, frequency, days, time, timezone, interval, interval_unit, start_date}}
//                                         -> scheduled processing (see _shared/schedule.ts)
//   DELETE ?id=<mailbox id>               -> disconnect (Vault secret removed by trigger)
//   GET ?code=...&state=...               -> OAuth callback (either provider); stores the refresh token for the
//                                            chosen number of days, redirects back to the app's dashboard
//                                            (or to connect-done, which reports to the opener, when signing in
//                                            in a popup)
// verify_jwt is off for this function (the OAuth callback carries no Supabase JWT); every other
// route checks the caller itself. Failures are written to app_logs (see _shared/log.ts).
import { decodeBase64Url, encodeBase64Url } from "jsr:@std/encoding@1/base64url";
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { GOOGLE_AUTH_URL, GOOGLE_SCOPES, GOOGLE_TOKEN_URL } from "../_shared/google.ts";
import { type Provider, PROVIDER_LABELS, PROVIDERS } from "../_shared/mail.ts";
import { type AppCredentials, appToken, mailboxPath, MS_SCOPES, MS_SHARED_SCOPES, msBase } from "../_shared/microsoft.ts";
import {
  FREQUENCIES, type Frequency, INTERVAL_UNITS, type IntervalUnit, isValidTimeZone, localDate, MIN_INTERVAL_MINUTES, nextRun,
} from "../_shared/schedule.ts";
import { adminClient, callerUser, env } from "../_shared/supabase.ts";
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";

const ACCESS_DAYS = [1, 7, 30, 90];
const DEFAULT_ACCESS_DAYS = 30;

const redirectUri = () => `${env("SUPABASE_URL")}/functions/v1/mailboxes`;

// The app origin travels in the signed OAuth state (taken from the browser's Origin header when connecting),
// so the callback returns to whichever deployment started it. APP_URL is only a fallback.
function appOrigin(req: Request): string | null {
  const origin = req.headers.get("Origin");
  if (origin && /^https?:\/\/[^/]+$/.test(origin)) return origin;
  return Deno.env.get("APP_URL")?.replace(/\/+$/, "") || null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  const url = new URL(req.url);

  if (req.method === "GET" && (url.searchParams.has("code") || url.searchParams.has("error"))) {
    return await oauthCallback(url);
  }

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;

  try {
    if (req.method === "GET") {
      const { error: expireError } = await adminClient().rpc("expire_mailbox_access");
      if (expireError) await logApi("mailboxes/expire", expireError.message, { userId: user.id, level: "warn" });

      const { data, error } = await sb
        .from("mailboxes")
        .select("id, provider, auth_type, email_address, delegate_address, display_name, active, smart_ap_address, last_processed_at, access_expires_at, created_at, schedule_enabled, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_interval, schedule_interval_unit, schedule_anchor, schedule_next_run_at, schedule_last_run_at, schedule_last_result")
        .order("created_at");
      if (error) return await fail("mailboxes/list", error.message, 500, { userId: user.id });
      return json({ mailboxes: data });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action === "connect_app") return await connectApp(user.id, body);
      if (body.action !== "connect") return json({ error: "Unknown action" }, 400);
      const provider: Provider = body.provider ?? "microsoft";
      if (!PROVIDERS.includes(provider)) return json({ error: "provider must be microsoft or google" }, 400);
      const missing = [...OAUTH[provider].settings, "OAUTH_STATE_SECRET"].filter((n) => !Deno.env.get(n));
      if (missing.length) {
        const name = provider === "google" ? "Gmail" : "Azure";
        return await fail("mailboxes/connect", `${name} isn't set up on the server yet (missing ${missing.join(", ")}).`, 400, { userId: user.id });
      }
      const origin = appOrigin(req);
      if (!origin) return await fail("mailboxes/connect", "Could not tell which site to return to (no Origin header or APP_URL).", 400, { userId: user.id });
      const days = ACCESS_DAYS.includes(Number(body.access_days)) ? Number(body.access_days) : DEFAULT_ACCESS_DAYS;
      const shared = String(body.shared_mailbox ?? "").trim().toLowerCase() || undefined;
      if (shared && provider !== "microsoft") return json({ error: "Shared mailboxes are only supported for Azure / Microsoft 365" }, 400);
      if (shared && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(shared)) return json({ error: "Shared mailbox must be a valid email address" }, 400);
      const s: State = { uid: user.id, days, origin, popup: body.popup === true, provider, shared };
      return json({ url: OAUTH[provider].authorizeUrl(await signState(s), s) });
    }

    if (req.method === "PATCH") {
      const body = await req.json().catch(() => ({}));
      if (body.id && body.schedule) return await saveSchedule(user.id, sb, body.id, body.schedule);
      const address = String(body.smart_ap_address ?? "").trim().toLowerCase();
      if (!body.id) return json({ error: "id is required" }, 400);
      if (address && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) {
        return json({ error: "Smart AP address must be a valid email address" }, 400);
      }
      const { data, error } = await sb.from("mailboxes")
        .update({ smart_ap_address: address || null, updated_at: new Date().toISOString() })
        .eq("id", body.id)
        .select("id, smart_ap_address")
        .maybeSingle();
      if (error) return await fail("mailboxes/update", error.message, 500, { userId: user.id, context: { id: body.id } });
      if (!data) return json({ error: "Mailbox not found" }, 404);
      return json({ mailbox: data });
    }

    if (req.method === "DELETE") {
      const id = url.searchParams.get("id");
      if (!id) return json({ error: "id is required" }, 400);
      const { error, count } = await sb.from("mailboxes").delete({ count: "exact" }).eq("id", id);
      if (error) return await fail("mailboxes/delete", error.message, 500, { userId: user.id, context: { id } });
      if (!count) return json({ error: "Mailbox not found" }, 404);
      return json({ deleted: id });
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("mailboxes", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

// Saves a mailbox's processing schedule and computes its next run. Schedule columns are written with the
// service role (after checking ownership) so users can't set next/last run times directly.
async function saveSchedule(userId: string, sb: SupabaseClient, id: string, s: Record<string, unknown>): Promise<Response> {
  const { data: owned, error: ownError } = await sb.from("mailboxes").select("id").eq("id", id).maybeSingle();
  if (ownError) return await fail("mailboxes/schedule", ownError.message, 500, { userId });
  if (!owned) return json({ error: "Mailbox not found" }, 404);

  const enabled = s.enabled === true;
  const frequency = String(s.frequency) as Frequency;
  const days = Array.isArray(s.days) ? [...new Set(s.days.map(Number).filter((d) => Number.isInteger(d) && d >= 0 && d <= 6))].sort() : [];
  const time = String(s.time ?? "");
  const timezone = String(s.timezone ?? "");
  const interval = Number(s.interval ?? 1);
  const intervalUnit = String(s.interval_unit ?? "hours") as IntervalUnit;
  if (!FREQUENCIES.includes(frequency)) return json({ error: "frequency must be daily, alternate, weekdays or interval" }, 400);
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) return json({ error: "time must be HH:MM" }, 400);
  if (!isValidTimeZone(timezone)) return json({ error: "Unknown time zone" }, 400);
  if (frequency === "weekdays" && days.length === 0) return json({ error: "Select at least one day" }, 400);
  if (frequency === "interval") {
    if (!INTERVAL_UNITS.includes(intervalUnit)) return json({ error: "interval_unit must be minutes, hours or days" }, 400);
    if (!Number.isInteger(interval) || interval < 1 || interval > 1000) return json({ error: "interval must be a whole number from 1 to 1000" }, 400);
    if (intervalUnit === "minutes" && interval < MIN_INTERVAL_MINUTES) {
      return json({ error: `Minute intervals must be at least ${MIN_INTERVAL_MINUTES} minutes` }, 400);
    }
  }
  const startDate = String(s.start_date ?? "");
  if (frequency === "interval" && !/^\d{4}-\d{2}-\d{2}$/.test(startDate)) return json({ error: "start_date must be YYYY-MM-DD" }, 400);

  const now = new Date();
  // Interval schedules count from the chosen start date; the others from the day the schedule is saved.
  const anchor = frequency === "interval" ? startDate : localDate(now, timezone);
  const next = enabled ? nextRun({ frequency, days, time, timezone, anchor, interval, intervalUnit }, now) : null;

  const { data, error } = await adminClient().from("mailboxes").update({
    schedule_enabled: enabled,
    schedule_frequency: frequency,
    schedule_days: days.length ? days : [1, 2, 3, 4, 5],
    schedule_time: time,
    schedule_timezone: timezone,
    schedule_anchor: anchor,
    ...(frequency === "interval" ? { schedule_interval: interval, schedule_interval_unit: intervalUnit } : {}),
    schedule_next_run_at: next?.toISOString() ?? null,
    // Next scheduled run covers mail since now, not everything since the mailbox was connected.
    ...(enabled ? { schedule_last_run_at: now.toISOString() } : {}),
    updated_at: now.toISOString(),
  }).eq("id", id).select("id, schedule_enabled, schedule_next_run_at").single();
  if (error) return await fail("mailboxes/schedule", error.message, 500, { userId, context: { id } });
  return json({ mailbox: data });
}

// Always ends with a redirect to the app's dashboard (connect-done in a popup); the outcome travels in
// ?connected= / ?connect_error=.
async function oauthCallback(url: URL): Promise<Response> {
  const state = await verifyState(url.searchParams.get("state") ?? "").catch(() => null);
  const userId = state?.uid ?? null;

  const back = (params: Record<string, string>) => {
    const origin = state?.origin ?? Deno.env.get("APP_URL")?.replace(/\/+$/, "");
    if (!origin) return json({ error: params.connect_error ?? "Connected. Return to the app." }, 400);
    return Response.redirect(`${origin}/${state?.popup ? "connect-done" : "dashboard"}?${new URLSearchParams(params)}`, 302);
  };
  const failed = async (message: string, err: unknown = message, context?: Record<string, unknown>) => {
    await logApi("mailboxes/callback", err, { userId, context });
    return back({ connect_error: message });
  };

  const label = PROVIDER_LABELS[state?.provider ?? "microsoft"];
  if (url.searchParams.get("error")) {
    const detail = url.searchParams.get("error_description") ?? url.searchParams.get("error")!;
    return await failed(detail, `${label} returned an error: ${detail}`, { error: url.searchParams.get("error") });
  }
  if (!state) return await failed("Sign-in expired or was tampered with. Please try again.");

  try {
    const { refreshToken, email, displayName, delegate } = await OAUTH[state.provider].exchange(url.searchParams.get("code")!, state);

    await saveMailbox(state.uid, {
      provider: state.provider, auth_type: "oauth", email_address: email, delegate_address: delegate, display_name: displayName,
    }, refreshToken, state.days);
    return back({ connected: email, days: String(state.days) });
  } catch (err) {
    if (err instanceof ConnectError) return await failed(err.message, err.detail, err.context);
    const message = err instanceof Error ? err.message : String(err);
    return await failed(`Could not save the mailbox: ${message}`, err);
  }
}

// Creates or updates the user's mailbox row and stores its credential (refresh token or API keys) in Vault for
// `days` days.
async function saveMailbox(
  uid: string,
  fields: { provider: Provider; auth_type: "oauth" | "app"; email_address: string; delegate_address: string | null; display_name: string | null },
  secret: string,
  days: number,
): Promise<string> {
  const db = adminClient();
  const { data: existing, error: findError } = await db.from("mailboxes").select("id, credentials_secret_id")
    .eq("owner_user_id", uid).eq("email_address", fields.email_address).maybeSingle();
  if (findError) throw new Error(findError.message);

  let id = existing?.id as string | undefined;
  if (!id) {
    const { data, error } = await db.from("mailboxes").insert({ owner_user_id: uid, ...fields }).select("id").single();
    if (error) throw new Error(error.message);
    id = data.id as string;
  }

  const { data: secretId, error: secretError } = await db.rpc("mailbox_secret_put", {
    p_secret_id: existing?.credentials_secret_id ?? null,
    p_secret: secret,
    p_name: `mailbox:${id}`,
  });
  if (secretError) throw new Error(`Could not store mailbox credentials: ${secretError.message}`);

  const { error } = await db.from("mailboxes").update({
    ...fields,
    credentials_secret_id: secretId,
    active: true,
    access_expires_at: new Date(Date.now() + days * 86_400_000).toISOString(),
    updated_at: new Date().toISOString(),
  }).eq("id", id);
  if (error) throw new Error(error.message);
  return id;
}

// Connects a Microsoft mailbox with API keys after checking they work: a token, then reading the inbox. Sending
// is not checked here; without Mail.Send, forwarding and the summary email fail and are logged.
async function connectApp(userId: string, body: Record<string, unknown>): Promise<Response> {
  const email = String(body.email_address ?? "").trim().toLowerCase();
  const creds: AppCredentials = {
    tenant_id: String(body.tenant_id ?? "").trim(),
    client_id: String(body.client_id ?? "").trim(),
    client_secret: String(body.client_secret ?? "").trim(),
  };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ error: "Mailbox must be a valid email address" }, 400);
  if (!creds.tenant_id || !creds.client_id || !creds.client_secret) {
    return json({ error: "Tenant ID, client ID and client secret are required" }, 400);
  }
  const days = ACCESS_DAYS.includes(Number(body.access_days)) ? Number(body.access_days) : DEFAULT_ACCESS_DAYS;
  // Never log the secret.
  const context = { mailbox: email, tenant_id: creds.tenant_id, client_id: creds.client_id };

  let token: string;
  try {
    token = await appToken(creds);
  } catch (err) {
    return await fail("mailboxes/connect_app", err instanceof Error ? err.message : String(err), 400, { userId, context });
  }
  const res = await fetch(`https://graph.microsoft.com/v1.0${mailboxPath(email)}/mailFolders/inbox?$select=id`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const message = res.status === 403 || res.status === 401
      ? `The app can't read ${email}. Grant it the Mail.Read application permission with admin consent (and check any Exchange access policy that limits which mailboxes it may open).`
      : res.status === 404 ? `No mailbox ${email} was found in that tenant.` : `Could not open ${email} (Graph ${res.status}).`;
    return await fail("mailboxes/connect_app", message, 400, { userId, context: { ...context, response: (await res.text()).slice(0, 2000) } });
  }

  const id = await saveMailbox(userId, {
    provider: "microsoft", auth_type: "app", email_address: email, delegate_address: null, display_name: null,
  }, JSON.stringify(creds), days);
  return json({ mailbox: { id, email_address: email }, days });
}

// --- per-provider OAuth: where sign-in starts, and how the callback's code becomes a refresh token ----------

// email is the mailbox to read; delegate is the signed-in account when that is a shared mailbox.
interface Identity { refreshToken: string; email: string; displayName: string | null; delegate: string | null }

// Thrown during the code exchange with the message shown to the user; `detail` and `context` go to the logs.
class ConnectError extends Error {
  constructor(message: string, readonly detail: unknown = message, readonly context?: Record<string, unknown>) {
    super(message);
  }
}

const OAUTH: Record<Provider, { settings: string[]; authorizeUrl(token: string, s: State): string; exchange(code: string, s: State): Promise<Identity> }> = {
  microsoft: {
    settings: ["MS_CLIENT_ID", "MS_CLIENT_SECRET"],
    authorizeUrl: (state, s) => `${msBase()}/authorize?` + new URLSearchParams({
      client_id: env("MS_CLIENT_ID"),
      response_type: "code",
      redirect_uri: redirectUri(),
      response_mode: "query",
      scope: s.shared ? MS_SHARED_SCOPES : MS_SCOPES,
      prompt: "select_account",
      state,
    }),
    async exchange(code, s) {
      const tokenRes = await fetch(`${msBase()}/token`, {
        method: "POST",
        body: new URLSearchParams({
          client_id: env("MS_CLIENT_ID"),
          client_secret: env("MS_CLIENT_SECRET"),
          grant_type: "authorization_code",
          code,
          redirect_uri: redirectUri(),
          scope: s.shared ? MS_SHARED_SCOPES : MS_SCOPES,
        }),
      });
      if (!tokenRes.ok) {
        throw new ConnectError("Microsoft sign-in failed. Please try again.", `Token exchange failed: ${tokenRes.status}`, {
          response: (await tokenRes.text()).slice(0, 2000),
        });
      }
      const tok = await tokenRes.json();
      if (!tok.refresh_token) {
        throw new ConnectError("Microsoft didn't return a refresh token. Remove the app's access and connect again.");
      }

      const meRes = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName", {
        headers: { Authorization: `Bearer ${tok.access_token}` },
      });
      if (!meRes.ok) {
        throw new ConnectError("Could not read the mailbox address.", `Graph /me failed: ${meRes.status}`, {
          response: (await meRes.text()).slice(0, 2000),
        });
      }
      const me = await meRes.json();
      const email = String(me.mail ?? me.userPrincipalName ?? "").toLowerCase();
      if (!email) throw new ConnectError("Could not read the mailbox address.");
      if (!s.shared || s.shared === email) return { refreshToken: tok.refresh_token, email, displayName: me.displayName ?? null, delegate: null };

      // Check the signed-in account can actually open the shared mailbox before saving it.
      const boxRes = await fetch(`https://graph.microsoft.com/v1.0${mailboxPath(s.shared)}/mailFolders/inbox?$select=id`, {
        headers: { Authorization: `Bearer ${tok.access_token}` },
      });
      if (!boxRes.ok) {
        throw new ConnectError(
          `${email} can't open the shared mailbox ${s.shared}. Ask your Microsoft 365 admin for Full Access to it, or check the address.`,
          `Graph shared mailbox check failed: ${boxRes.status}`, { shared: s.shared, response: (await boxRes.text()).slice(0, 2000) },
        );
      }
      return { refreshToken: tok.refresh_token, email: s.shared, displayName: null, delegate: email };
    },
  },

  google: {
    settings: ["GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET"],
    // access_type=offline + prompt=consent so Google returns a refresh token on every connect.
    authorizeUrl: (state) => `${GOOGLE_AUTH_URL}?` + new URLSearchParams({
      client_id: env("GOOGLE_CLIENT_ID"),
      response_type: "code",
      redirect_uri: redirectUri(),
      scope: GOOGLE_SCOPES,
      access_type: "offline",
      prompt: "consent select_account",
      state,
    }),
    async exchange(code) {
      const tokenRes = await fetch(GOOGLE_TOKEN_URL, {
        method: "POST",
        body: new URLSearchParams({
          client_id: env("GOOGLE_CLIENT_ID"),
          client_secret: env("GOOGLE_CLIENT_SECRET"),
          grant_type: "authorization_code",
          code,
          redirect_uri: redirectUri(),
        }),
      });
      if (!tokenRes.ok) {
        throw new ConnectError("Google sign-in failed. Please try again.", `Token exchange failed: ${tokenRes.status}`, {
          response: (await tokenRes.text()).slice(0, 2000),
        });
      }
      const tok = await tokenRes.json();
      if (!tok.refresh_token) {
        throw new ConnectError("Google didn't return a refresh token. Remove the app's access in your Google account and connect again.");
      }
      // Google lets people untick individual permissions on the consent screen.
      const granted = String(tok.scope ?? "").split(" ");
      if (!GOOGLE_SCOPES.split(" ").every((sc) => granted.includes(sc))) {
        throw new ConnectError("Allow SummarizeMail to read and send Gmail on Google's consent screen, then try again.",
          "Google consent missing Gmail scopes", { granted });
      }

      const profileRes = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/profile", {
        headers: { Authorization: `Bearer ${tok.access_token}` },
      });
      if (!profileRes.ok) {
        throw new ConnectError("Could not read the mailbox address.", `Gmail profile failed: ${profileRes.status}`, {
          response: (await profileRes.text()).slice(0, 2000),
        });
      }
      const email = String((await profileRes.json()).emailAddress ?? "").toLowerCase();
      if (!email) throw new ConnectError("Could not read the mailbox address.");
      return { refreshToken: tok.refresh_token, email, displayName: null, delegate: null };
    },
  },
};

// --- signed OAuth state (HMAC-SHA256, 15 minute expiry) ----------------------------------

interface State {
  uid: string;
  days: number;
  origin: string;
  popup?: boolean;
  provider: Provider;
  // Microsoft shared mailbox address to read through the signed-in account.
  shared?: string;
}

const enc = new TextEncoder();

async function hmacKey() {
  return await crypto.subtle.importKey("raw", enc.encode(env("OAUTH_STATE_SECRET")),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
}

async function signState(s: State): Promise<string> {
  const payload = encodeBase64Url(enc.encode(JSON.stringify({ ...s, exp: Date.now() + 15 * 60_000 })));
  const sig = new Uint8Array(await crypto.subtle.sign("HMAC", await hmacKey(), enc.encode(payload)));
  return `${payload}.${encodeBase64Url(sig)}`;
}

async function verifyState(token: string): Promise<State | null> {
  const [payload, sig] = token.split(".");
  if (!payload || !sig) return null;
  const ok = await crypto.subtle.verify("HMAC", await hmacKey(), decodeBase64Url(sig), enc.encode(payload));
  if (!ok) return null;
  const s = JSON.parse(new TextDecoder().decode(decodeBase64Url(payload)));
  if (!(s.exp > Date.now()) || typeof s.uid !== "string" || typeof s.origin !== "string") return null;
  return { uid: s.uid, days: ACCESS_DAYS.includes(s.days) ? s.days : DEFAULT_ACCESS_DAYS, origin: s.origin, popup: s.popup === true,
    provider: PROVIDERS.includes(s.provider) ? s.provider : "microsoft", shared: typeof s.shared === "string" ? s.shared : undefined };
}
