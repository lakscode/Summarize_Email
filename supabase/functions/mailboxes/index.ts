// Azure / Microsoft 365 mailboxes for the signed-in user.
//   GET                                   -> { mailboxes: [...] } (expired access is cleaned up first)
//   POST {action: "connect", access_days} -> { url } to start Microsoft sign-in (delegated OAuth)
//   PATCH {id, smart_ap_address}          -> set where invoices are forwarded ("" clears it)
//   PATCH {id, schedule: {enabled, frequency, days, time, timezone, interval, interval_unit, start_date}}
//                                         -> scheduled processing (see _shared/schedule.ts)
//   DELETE ?id=<mailbox id>               -> disconnect (Vault secret removed by trigger)
//   GET ?code=...&state=...               -> OAuth callback from Microsoft; stores the refresh token for the
//                                            chosen number of days, redirects back to the app's dashboard
// verify_jwt is off for this function (the Microsoft callback carries no Supabase JWT); every other
// route checks the caller itself. Failures are written to app_logs (see _shared/log.ts).
import { decodeBase64Url, encodeBase64Url } from "jsr:@std/encoding@1/base64url";
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { MS_SCOPES, msBase } from "../_shared/microsoft.ts";
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
        .select("id, provider, email_address, display_name, active, smart_ap_address, last_processed_at, access_expires_at, created_at, schedule_enabled, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_interval, schedule_interval_unit, schedule_anchor, schedule_next_run_at, schedule_last_run_at, schedule_last_result")
        .order("created_at");
      if (error) return await fail("mailboxes/list", error.message, 500, { userId: user.id });
      return json({ mailboxes: data });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action !== "connect") return json({ error: "Unknown action" }, 400);
      const missing = ["MS_CLIENT_ID", "MS_CLIENT_SECRET", "OAUTH_STATE_SECRET"].filter((n) => !Deno.env.get(n));
      if (missing.length) {
        return await fail("mailboxes/connect", `Azure isn't set up on the server yet (missing ${missing.join(", ")}).`, 400, { userId: user.id });
      }
      const origin = appOrigin(req);
      if (!origin) return await fail("mailboxes/connect", "Could not tell which site to return to (no Origin header or APP_URL).", 400, { userId: user.id });
      const days = ACCESS_DAYS.includes(Number(body.access_days)) ? Number(body.access_days) : DEFAULT_ACCESS_DAYS;

      return json({
        url: `${msBase()}/authorize?` + new URLSearchParams({
          client_id: env("MS_CLIENT_ID"),
          response_type: "code",
          redirect_uri: redirectUri(),
          response_mode: "query",
          scope: MS_SCOPES,
          prompt: "select_account",
          state: await signState({ uid: user.id, days, origin }),
        }),
      });
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

// Always ends with a redirect to the app's dashboard; the outcome travels in ?connected= / ?connect_error=.
async function oauthCallback(url: URL): Promise<Response> {
  const state = await verifyState(url.searchParams.get("state") ?? "").catch(() => null);
  const userId = state?.uid ?? null;

  const back = (params: Record<string, string>) => {
    const origin = state?.origin ?? Deno.env.get("APP_URL")?.replace(/\/+$/, "");
    if (!origin) return json({ error: params.connect_error ?? "Connected. Return to the app." }, 400);
    return Response.redirect(`${origin}/dashboard?${new URLSearchParams(params)}`, 302);
  };
  const failed = async (message: string, err: unknown = message, context?: Record<string, unknown>) => {
    await logApi("mailboxes/callback", err, { userId, context });
    return back({ connect_error: message });
  };

  if (url.searchParams.get("error")) {
    const detail = url.searchParams.get("error_description") ?? url.searchParams.get("error")!;
    return await failed(detail, `Microsoft returned an error: ${detail}`, { error: url.searchParams.get("error") });
  }
  if (!state) return await failed("Sign-in expired or was tampered with. Please try again.");

  try {
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
      const body = await tokenRes.text();
      return await failed("Microsoft sign-in failed. Please try again.", `Token exchange failed: ${tokenRes.status}`, {
        response: body.slice(0, 2000),
      });
    }
    const tok = await tokenRes.json();
    if (!tok.refresh_token) {
      return await failed("Microsoft didn't return a refresh token. Remove the app's access and connect again.");
    }

    const meRes = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName", {
      headers: { Authorization: `Bearer ${tok.access_token}` },
    });
    if (!meRes.ok) {
      return await failed("Could not read the mailbox address.", `Graph /me failed: ${meRes.status}`, {
        response: (await meRes.text()).slice(0, 2000),
      });
    }
    const me = await meRes.json();
    const email = String(me.mail ?? me.userPrincipalName ?? "").toLowerCase();
    if (!email) return await failed("Could not read the mailbox address.");

    const db = adminClient();
    const { data: existing, error: findError } = await db.from("mailboxes").select("id, credentials_secret_id")
      .eq("owner_user_id", state.uid).eq("email_address", email).maybeSingle();
    if (findError) throw new Error(findError.message);

    let id = existing?.id as string | undefined;
    if (!id) {
      const { data, error } = await db.from("mailboxes")
        .insert({ owner_user_id: state.uid, provider: "microsoft", email_address: email, display_name: me.displayName ?? null })
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

    const expiresAt = new Date(Date.now() + state.days * 86_400_000);
    const { error } = await db.from("mailboxes").update({
      credentials_secret_id: secretId,
      display_name: me.displayName ?? null,
      active: true,
      access_expires_at: expiresAt.toISOString(),
      updated_at: new Date().toISOString(),
    }).eq("id", id);
    if (error) throw new Error(error.message);

    return back({ connected: email, days: String(state.days) });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return await failed(`Could not save the mailbox: ${message}`, err);
  }
}

// --- signed OAuth state (HMAC-SHA256, 15 minute expiry) ----------------------------------

interface State {
  uid: string;
  days: number;
  origin: string;
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
  return { uid: s.uid, days: ACCESS_DAYS.includes(s.days) ? s.days : DEFAULT_ACCESS_DAYS, origin: s.origin };
}
