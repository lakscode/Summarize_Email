// Connected mailboxes (Azure / Microsoft 365 or Gmail) for the signed-in user.
//   GET                                   -> { mailboxes, connections, categories } (expired access is cleaned up
//                                            first; categories: the AP and AR categories per mailbox type)
//   POST {action: "connect", provider: "microsoft" | "google", access_days, popup, shared_mailbox?}
//                                         -> { url } to start Microsoft or Google sign-in (OAuth). With
//                                            shared_mailbox (Microsoft only), the user signs in with their own
//                                            account and the app reads that shared mailbox through it.
//   Organization connections (API keys): one Microsoft 365 tenant's Azure app registration (application permissions
//   Mail.Read and Mail.Send, admin-consented), entered once; its shared mailboxes are then added by address. Keys are
//   checked before saving, kept in Vault and never returned.
//   POST {action: "add_connection", name, tenant_id, client_id, client_secret}       -> { connection, warning? }
//   POST {action: "update_connection", connection_id, client_secret}                 -> { connection } (new secret)
//   POST {action: "add_mailboxes", connection_id, addresses, mailbox_type, smart_ap_address?, forward_categories?}
//                                         -> { results: [{address, ok, error?}] } each checked for inbox access
//   POST {action: "find_shared_mailboxes", connection_id, person}
//                                         -> { mailboxes: [{address, name}], checked, total } shared mailboxes the
//                                            person has Full Access to (Exchange admin API, see microsoft.ts)
//   DELETE ?connection_id=<id>            -> remove the connection, its mailboxes and their processed emails
//   PATCH {id, mailbox_type?, forward_categories?, smart_ap_address?}
//                                         -> per-mailbox settings: AP or AR, which categories are forwarded, and the
//                                            destination address they go to ("" clears it)
//   PATCH {id, schedule: {enabled, frequency, days, time, timezone, interval, interval_unit, start_date}}
//                                         -> scheduled processing (see _shared/schedule.ts)
//   DELETE ?id=<mailbox id>               -> disconnect (Vault secret removed by trigger)
//   GET ?code=...&state=...               -> OAuth callback (either provider); stores the refresh token for the
//                                            chosen number of days, redirects back to the app's dashboard
//                                            (or to connect-done, which reports to the opener, when signing in
//                                            in a popup)
// verify_jwt is off for this function (the OAuth callback carries no Supabase JWT); every other
// route checks the caller itself. Failures are written to app_logs (see _shared/log.ts).
import { decodeBase64Url, encodeBase64Url } from "../_shared/base64.ts";
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { GOOGLE_AUTH_URL, GOOGLE_SCOPES, GOOGLE_TOKEN_URL } from "../_shared/google.ts";
import { type Provider, PROVIDER_LABELS, PROVIDERS } from "../_shared/mail.ts";
import { CATEGORIES, DEFAULT_FORWARD, MAILBOX_TYPES, type MailboxType } from "../_shared/classify.ts";
import {
  type AppCredentials, appToken, connectionCredentials, mailboxPath, MS_SCOPES, MS_SHARED_SCOPES, msBase, sharedMailboxesFor,
} from "../_shared/microsoft.ts";
import {
  FREQUENCIES, type Frequency, INTERVAL_UNITS, type IntervalUnit, isValidTimeZone, localDate, MIN_INTERVAL_MINUTES, nextRun,
} from "../_shared/schedule.ts";
import { adminClient, callerUser, env } from "../_shared/supabase.ts";
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";

const ACCESS_DAYS = [1, 7, 30, 90];
const DEFAULT_ACCESS_DAYS = 30;

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// Mailboxes per add_mailboxes call, and how many are checked against Graph at once.
const MAX_BULK = 100;
const CHECK_CONCURRENCY = 5;

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
        .select("id, provider, auth_type, connection_id, mailbox_type, forward_categories, email_address, delegate_address, display_name, active, smart_ap_address, last_processed_at, access_expires_at, created_at, schedule_enabled, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_interval, schedule_interval_unit, schedule_anchor, schedule_next_run_at, schedule_last_run_at, schedule_last_result")
        .order("created_at");
      if (error) return await fail("mailboxes/list", error.message, 500, { userId: user.id });
      const { data: connections, error: connError } = await sb.from("ms_connections")
        .select("id, name, tenant_id, client_id, created_at").order("created_at");
      if (connError) return await fail("mailboxes/list", connError.message, 500, { userId: user.id });
      return json({ mailboxes: data, connections, categories: CATEGORIES });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action === "add_connection") return await addConnection(user.id, body);
      if (body.action === "update_connection") return await updateConnection(user.id, sb, body);
      if (body.action === "add_mailboxes") return await addMailboxes(user.id, sb, body);
      if (body.action === "find_shared_mailboxes") return await findSharedMailboxes(user.id, sb, body);
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
      if (shared && !EMAIL.test(shared)) return json({ error: "Shared mailbox must be a valid email address" }, 400);
      const s: State = { uid: user.id, days, origin, popup: body.popup === true, provider, shared };
      return json({ url: OAUTH[provider].authorizeUrl(await signState(s), s) });
    }

    if (req.method === "PATCH") {
      const body = await req.json().catch(() => ({}));
      if (body.id && body.schedule) return await saveSchedule(user.id, sb, body.id, body.schedule);
      if (!body.id) return json({ error: "id is required" }, 400);
      return await saveSettings(user.id, sb, body);
    }

    if (req.method === "DELETE" && url.searchParams.has("connection_id")) {
      const id = url.searchParams.get("connection_id")!;
      const { error, count } = await sb.from("ms_connections").delete({ count: "exact" }).eq("id", id);
      if (error) return await fail("mailboxes/delete_connection", error.message, 500, { userId: user.id, context: { id } });
      if (!count) return json({ error: "Connection not found" }, 404);
      return json({ deleted: id });
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
      provider: state.provider, auth_type: "oauth", connection_id: null, email_address: email, delegate_address: delegate,
      display_name: displayName,
    }, refreshToken, state.days);
    return back({ connected: email, days: String(state.days) });
  } catch (err) {
    if (err instanceof ConnectError) return await failed(err.message, err.detail, err.context);
    const message = err instanceof Error ? err.message : String(err);
    return await failed(`Could not save the mailbox: ${message}`, err);
  }
}

// Creates or updates the user's signed-in mailbox row and stores its refresh token in Vault for `days` days.
async function saveMailbox(
  uid: string,
  fields: {
    provider: Provider; auth_type: "oauth"; connection_id: null; email_address: string; delegate_address: string | null;
    display_name: string | null;
  },
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

// Per-mailbox settings. forward_categories must be categories of the mailbox's type; changing the type without
// sending forward_categories resets them to that type's default.
async function saveSettings(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const { data: current, error: readError } = await sb.from("mailboxes").select("id, mailbox_type").eq("id", body.id).maybeSingle();
  if (readError) return await fail("mailboxes/update", readError.message, 500, { userId, context: { id: body.id } });
  if (!current) return json({ error: "Mailbox not found" }, 404);

  const update: Record<string, unknown> = { updated_at: new Date().toISOString() };
  if (body.smart_ap_address !== undefined) {
    const address = String(body.smart_ap_address ?? "").trim().toLowerCase();
    if (address && !EMAIL.test(address)) return json({ error: "Destination must be a valid email address" }, 400);
    update.smart_ap_address = address || null;
  }
  let type = current.mailbox_type as MailboxType;
  if (body.mailbox_type !== undefined) {
    if (!MAILBOX_TYPES.includes(body.mailbox_type as MailboxType)) return json({ error: "mailbox_type must be ap or ar" }, 400);
    if (body.mailbox_type !== type && body.forward_categories === undefined) update.forward_categories = DEFAULT_FORWARD[body.mailbox_type as MailboxType];
    type = body.mailbox_type as MailboxType;
    update.mailbox_type = type;
  }
  if (body.forward_categories !== undefined) {
    const cats = forwardCategories(body.forward_categories, type);
    if (!cats) return json({ error: `forward_categories must be from: ${forwardable(type).join(", ")}` }, 400);
    update.forward_categories = cats;
  }

  const { data, error } = await sb.from("mailboxes").update(update).eq("id", body.id)
    .select("id, mailbox_type, forward_categories, smart_ap_address").maybeSingle();
  if (error) return await fail("mailboxes/update", error.message, 500, { userId, context: { id: body.id } });
  if (!data) return json({ error: "Mailbox not found" }, 404);
  return json({ mailbox: data });
}

const forwardable = (type: MailboxType) => CATEGORIES[type].filter((c) => c !== "other");

// Validated, de-duplicated forward categories for `type`, or null if any is not forwardable for that type.
function forwardCategories(value: unknown, type: MailboxType): string[] | null {
  if (!Array.isArray(value)) return null;
  const cats = [...new Set(value.map(String))];
  return cats.every((c) => forwardable(type).includes(c)) ? cats : null;
}

// --- organization connections (API keys) ------------------------------------------------------------------

// Gets a token with the keys and reads the application permissions it carries (the "roles" claim), so a missing
// admin consent is reported now rather than on the first run.
async function checkKeys(creds: AppCredentials): Promise<{ token: string; roles: string[] }> {
  const token = await appToken(creds);
  let roles: string[] = [];
  try {
    roles = JSON.parse(new TextDecoder().decode(decodeBase64Url(token.split(".")[1]))).roles ?? [];
  } catch {
    // Not a readable JWT; the mailbox checks still catch missing permissions.
  }
  return { token, roles };
}

function credsFrom(body: Record<string, unknown>): AppCredentials {
  return {
    tenant_id: String(body.tenant_id ?? "").trim(),
    client_id: String(body.client_id ?? "").trim(),
    client_secret: String(body.client_secret ?? "").trim(),
  };
}

const MISSING_READ = "The Azure app has no Mail.Read application permission with admin consent. In Azure: API permissions > " +
  "Add > Microsoft Graph > Application permissions > Mail.Read (and Mail.Send), then Grant admin consent.";

async function addConnection(userId: string, body: Record<string, unknown>): Promise<Response> {
  const creds = credsFrom(body);
  const name = String(body.name ?? "").trim() || creds.tenant_id;
  if (!creds.tenant_id || !creds.client_id || !creds.client_secret) {
    return json({ error: "Tenant ID, client ID and client secret are required" }, 400);
  }
  // Never log the secret.
  const context = { tenant_id: creds.tenant_id, client_id: creds.client_id };
  let roles: string[];
  try {
    ({ roles } = await checkKeys(creds));
  } catch (err) {
    return await fail("mailboxes/add_connection", err instanceof Error ? err.message : String(err), 400, { userId, context });
  }
  if (!roles.some((r) => r === "Mail.Read" || r === "Mail.ReadWrite")) {
    return await fail("mailboxes/add_connection", MISSING_READ, 400, { userId, context: { ...context, roles } });
  }

  const db = adminClient();
  const { data: row, error } = await db.from("ms_connections")
    .insert({ owner_user_id: userId, name, tenant_id: creds.tenant_id, client_id: creds.client_id })
    .select("id, name, tenant_id, client_id, created_at").single();
  if (error) return await fail("mailboxes/add_connection", error.message, 500, { userId, context });
  const { data: secretId, error: secretError } = await db.rpc("mailbox_secret_put", {
    p_secret_id: null, p_secret: JSON.stringify(creds), p_name: `ms_connection:${row.id}`,
  });
  if (secretError) {
    await db.from("ms_connections").delete().eq("id", row.id);
    return await fail("mailboxes/add_connection", `Could not store the API keys: ${secretError.message}`, 500, { userId, context });
  }
  await db.from("ms_connections").update({ credentials_secret_id: secretId }).eq("id", row.id);

  const warning = roles.includes("Mail.Send") ? undefined
    : "Connected, but the app has no Mail.Send application permission, so forwarding and summary emails will fail until it is granted.";
  return json({ connection: row, warning });
}

// Replaces a connection's client secret (Azure client secrets expire), after checking the new one works.
async function updateConnection(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const { data: conn, error } = await sb.from("ms_connections").select("id, tenant_id, client_id").eq("id", body.connection_id).maybeSingle();
  if (error) return await fail("mailboxes/update_connection", error.message, 500, { userId });
  if (!conn) return json({ error: "Connection not found" }, 404);
  const creds = credsFrom({ ...conn, client_secret: body.client_secret });
  if (!creds.client_secret) return json({ error: "client_secret is required" }, 400);
  try {
    const { roles } = await checkKeys(creds);
    if (!roles.some((r) => r === "Mail.Read" || r === "Mail.ReadWrite")) throw new Error(MISSING_READ);
  } catch (err) {
    return await fail("mailboxes/update_connection", err instanceof Error ? err.message : String(err), 400,
      { userId, context: { id: conn.id, tenant_id: conn.tenant_id, client_id: conn.client_id } });
  }

  const db = adminClient();
  const { data: row } = await db.from("ms_connections").select("credentials_secret_id").eq("id", conn.id).single();
  const { data: secretId, error: secretError } = await db.rpc("mailbox_secret_put", {
    p_secret_id: row?.credentials_secret_id ?? null, p_secret: JSON.stringify(creds), p_name: `ms_connection:${conn.id}`,
  });
  if (secretError) return await fail("mailboxes/update_connection", `Could not store the API keys: ${secretError.message}`, 500, { userId });
  await db.from("ms_connections").update({ credentials_secret_id: secretId, updated_at: new Date().toISOString() }).eq("id", conn.id);
  return json({ connection: { id: conn.id } });
}

// Lists the shared mailboxes in the connection's tenant that `person` has Full Access to, for the UI to add in bulk.
async function findSharedMailboxes(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const { data: conn, error } = await sb.from("ms_connections").select("id").eq("id", body.connection_id).maybeSingle();
  if (error) return await fail("mailboxes/find_shared", error.message, 500, { userId });
  if (!conn) return json({ error: "Connection not found" }, 404);
  const person = String(body.person ?? "").trim().toLowerCase();
  if (!EMAIL.test(person)) return json({ error: "Enter the person's email address" }, 400);
  try {
    return json(await sharedMailboxesFor(await connectionCredentials(adminClient(), conn.id), person));
  } catch (err) {
    return await fail("mailboxes/find_shared", err instanceof Error ? err.message : String(err), 400, { userId, context: { connection: conn.id, person } });
  }
}

// Adds shared mailboxes under a connection, checking each can be read. `addresses` is an array or text with one
// address per line / comma / semicolon. A mailbox already connected another way is switched to the connection,
// keeping its history and settings unless new ones are given.
async function addMailboxes(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const { data: conn, error } = await sb.from("ms_connections").select("id").eq("id", body.connection_id).maybeSingle();
  if (error) return await fail("mailboxes/add_mailboxes", error.message, 500, { userId });
  if (!conn) return json({ error: "Connection not found" }, 404);

  const raw = Array.isArray(body.addresses) ? body.addresses.join("\n") : String(body.addresses ?? "");
  const addresses = [...new Set(raw.split(/[\s,;]+/).map((a) => a.trim().toLowerCase()).filter(Boolean))];
  if (!addresses.length) return json({ error: "Enter at least one mailbox address" }, 400);
  if (addresses.length > MAX_BULK) return json({ error: `Add at most ${MAX_BULK} mailboxes at a time` }, 400);

  const type = (body.mailbox_type ?? "ap") as MailboxType;
  if (!MAILBOX_TYPES.includes(type)) return json({ error: "mailbox_type must be ap or ar" }, 400);
  const destination = String(body.smart_ap_address ?? "").trim().toLowerCase();
  if (destination && !EMAIL.test(destination)) return json({ error: "Destination must be a valid email address" }, 400);
  const forward = body.forward_categories === undefined ? DEFAULT_FORWARD[type] : forwardCategories(body.forward_categories, type);
  if (!forward) return json({ error: `forward_categories must be from: ${forwardable(type).join(", ")}` }, 400);

  const db = adminClient();
  let token: string;
  try {
    ({ token } = await checkKeys(await connectionCredentials(db, conn.id)));
  } catch (err) {
    return await fail("mailboxes/add_mailboxes", err instanceof Error ? err.message : String(err), 400, { userId, context: { connection: conn.id } });
  }

  const addOne = async (address: string): Promise<{ address: string; ok: boolean; error?: string }> => {
    if (!EMAIL.test(address)) return { address, ok: false, error: "Not a valid email address" };
    const res = await fetch(`https://graph.microsoft.com/v1.0${mailboxPath(address)}/mailFolders/inbox?$select=id`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      const detail = (await res.text()).slice(0, 1000);
      await logApi("mailboxes/add_mailboxes", `Graph ${res.status} for ${address}`, { userId, level: "warn", context: { connection: conn.id, response: detail } });
      return {
        address, ok: false,
        error: res.status === 404 ? "No such mailbox in this tenant"
          : res.status === 401 || res.status === 403 ? "The app may not read this mailbox (Mail.Read consent, or an Exchange access policy excludes it)"
          : `Graph error ${res.status}`,
      };
    }

    const { data: existing, error: findError } = await db.from("mailboxes").select("id, credentials_secret_id")
      .eq("owner_user_id", userId).eq("email_address", address).maybeSingle();
    if (findError) return { address, ok: false, error: findError.message };
    const fields = {
      provider: "microsoft", auth_type: "app", connection_id: conn.id, delegate_address: null, active: true,
      credentials_secret_id: null, access_expires_at: null, mailbox_type: type, forward_categories: forward,
      ...(destination ? { smart_ap_address: destination } : {}), updated_at: new Date().toISOString(),
    };
    const { error: saveError } = existing
      ? await db.from("mailboxes").update(fields).eq("id", existing.id)
      : await db.from("mailboxes").insert({ owner_user_id: userId, email_address: address, ...fields });
    if (saveError) return { address, ok: false, error: saveError.message };
    // Its own sign-in token is no longer used.
    if (existing?.credentials_secret_id) await db.rpc("mailbox_secret_delete", { p_secret_id: existing.credentials_secret_id });
    return { address, ok: true };
  };

  const results: { address: string; ok: boolean; error?: string }[] = [];
  const queue = [...addresses];
  await Promise.all(Array.from({ length: Math.min(CHECK_CONCURRENCY, queue.length) }, async () => {
    for (let a = queue.shift(); a; a = queue.shift()) results.push(await addOne(a));
  }));
  results.sort((a, b) => addresses.indexOf(a.address) - addresses.indexOf(b.address));
  return json({ results });
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
