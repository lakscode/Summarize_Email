// Connected mailboxes (Azure / Microsoft 365 or Gmail) for the signed-in user, and their organization's mailboxes.
//   GET                                   -> { mailboxes, connections, categories, organization } (expired access is
//                                            cleaned up first; categories: the AP and AR categories per mailbox type;
//                                            each mailbox has the caller's `role` and its `owner_name`; organization:
//                                            the caller's {id, name, role} or null)
//   Organization mailboxes (see migrations/*_organization_mailboxes.sql): the owner shares a mailbox with their
//   organization; everyone in it then has a role on it: admin (settings, schedule, roles) or member (the default).
//   PATCH {id, org_shared: true | false}  -> owner: { mailbox } shared with their organization, or not any more
//   GET ?access=<mailbox id>              -> { people: [{user_id, email, full_name, role, fixed, you}], manage }
//                                            everyone in the mailbox's organization with their role on it
//   POST {action: "set_mailbox_role", mailbox_id, user_id, role: "admin" | "member"} -> mailbox admins: { member }
//   POST {action: "connect", provider: "microsoft" | "google", access_days, popup, shared_mailbox?, login_hint?}
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
//   Shared mailboxes of a signed-in Microsoft account (after "Add Azure mailbox", the callback returns mailbox_id):
//   POST {action: "discover_shared", mailbox_id}
//                                         -> { mailboxes: [{address, name}], warning? } shared mailboxes that account
//                                            has Full Access to (Autodiscover, see microsoft.ts), not yet added
//   POST {action: "add_shared", mailbox_id, addresses}
//                                         -> { results: [{address, ok, error?}] } each added as a shared mailbox read
//                                            through that account (its token, same access period), after checking
//                                            the account can open it
//   POST {action: "auto_shared", mailbox_id}
//                                         -> { results: [{address, ok, error?}], warning? } discover_shared + add_shared
//                                            for every shared mailbox found, run after signing in to the app with
//                                            Microsoft so nothing has to be picked by hand
//   Shared mailboxes (add_shared, auto_shared, a shared mailbox sign-in, add_mailboxes): one copy in the whole app,
//   owned by whoever added it first and always shared with everyone in their organization; adding one that anyone
//   already has is refused (see migrations/*_one_shared_mailbox_per_organization.sql).
//   A newly added mailbox that another member of the same organization already has takes that member's settings
//   (type, forwarded categories, destination, summary); the schedule is not copied, so the same mailbox isn't
//   processed twice.
//   PATCH {id, mailbox_type?, forward_categories?, smart_ap_address?, send_summary?, summary_address?}
//                                         -> per-mailbox settings: AP or AR, which categories are forwarded, the
//                                            destination address they go to ("" clears it), whether the run summary
//                                            is emailed, and to which addresses (comma-separated; "" clears them: an
//                                            own mailbox gets it itself, a shared mailbox sends it to the owner's
//                                            account email)
//   PATCH {id, forward_report: {enabled, time, timezone, addresses}}
//                                         -> daily email listing the invoices forwarded from the mailbox, at `time`
//                                            ("HH:MM" in `timezone`) to `addresses` (one or more; see
//                                            _shared/forward-report.ts)
//   POST {action: "send_forward_report", mailbox_id} -> { result } that report now, for the time since the last one
//   PATCH {id, weekly_report: {enabled, day, time, timezone, addresses}}
//                                         -> weekly email listing the invoices forwarded in the past 7 days, on `day`
//                                            (0 = Sunday) at `time` to `addresses` (see _shared/weekly-report.ts)
//   POST {action: "send_weekly_report", mailbox_id} -> { result } the weekly report now, for the past 7 days
//   PATCH {id, schedule: {enabled, frequency, days, time, timezone, interval, interval_unit, start_date}}
//                                         -> scheduled processing (see _shared/schedule.ts)
//   Settings and schedule: the owner or a mailbox admin.
//   DELETE ?id=<mailbox id>               -> owner: disconnect (Vault secret removed by trigger)
//   GET ?code=...&state=...               -> OAuth callback (either provider); stores the refresh token for the
//                                            chosen number of days, redirects back to the app's dashboard
//                                            (or to connect-done, which reports to the opener, when signing in
//                                            in a popup)
// verify_jwt is off for this function (the OAuth callback carries no Supabase JWT); every other
// route checks the caller itself. Failures are written to app_logs (see _shared/log.ts).
import { decodeBase64Url, encodeBase64Url } from "../_shared/base64.ts";
import { APP_NAME } from "../_shared/brand.ts";
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { GOOGLE_AUTH_URL, GOOGLE_SCOPES, GOOGLE_TOKEN_URL } from "../_shared/google.ts";
import { type MailboxRow, type Provider, PROVIDER_LABELS, PROVIDERS, readRefreshToken } from "../_shared/mail.ts";
import { CATEGORIES, DEFAULT_FORWARD, MAILBOX_TYPES, type MailboxType } from "../_shared/classify.ts";
import {
  alternateMailboxes, type AppCredentials, appToken, connectionCredentials, delegatedToken, mailboxPath, MS_SHARED_SCOPES, msBase,
  openInbox, sharedMailboxesFor,
} from "../_shared/microsoft.ts";
import {
  FREQUENCIES, type Frequency, INTERVAL_UNITS, type IntervalUnit, isValidTimeZone, localDate, MIN_INTERVAL_MINUTES, nextRun,
} from "../_shared/schedule.ts";
import { nextReportAt, REPORT_MAILBOX_COLUMNS, type ReportMailbox, sendForwardReport } from "../_shared/forward-report.ts";
import {
  dayName, nextWeeklyReportAt, sendWeeklyReport, WEEKLY_REPORT_MAILBOX_COLUMNS, type WeeklyReportMailbox,
} from "../_shared/weekly-report.ts";
import { adminClient, callerUser, env } from "../_shared/supabase.ts";
import type { SupabaseClient } from "jsr:@supabase/supabase-js@2";

const ACCESS_DAYS = [1, 7, 30, 90];
const DEFAULT_ACCESS_DAYS = 30;

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// Addresses the run summary or the forwarded-invoices report may go to.
const MAX_SUMMARY_RECIPIENTS = 20;
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
    if (req.method === "GET" && url.searchParams.has("access")) return await mailboxAccess(user.id, url.searchParams.get("access"));

    if (req.method === "GET") {
      const { error: expireError } = await adminClient().rpc("expire_mailbox_access");
      if (expireError) await logApi("mailboxes/expire", expireError.message, { userId: user.id, level: "warn" });

      // RLS returns the caller's own mailboxes and those shared with their organization.
      const { data, error } = await sb
        .from("mailboxes")
        .select("id, owner_user_id, org_id, provider, auth_type, connection_id, mailbox_type, forward_categories, email_address, delegate_address, display_name, active, smart_ap_address, send_summary, summary_address, last_processed_at, last_processed_by_name, access_expires_at, created_at, schedule_enabled, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_interval, schedule_interval_unit, schedule_anchor, schedule_next_run_at, schedule_last_run_at, schedule_last_result, forward_report_enabled, forward_report_time, forward_report_timezone, forward_report_addresses, forward_report_next_at, forward_report_last_at, forward_report_last_result")
        .order("created_at");
      if (error) return await fail("mailboxes/list", error.message, 500, { userId: user.id });
      await addWeeklyReportSettings(user.id, sb, data);
      const { data: connections, error: connError } = await sb.from("ms_connections")
        .select("id, name, tenant_id, client_id, created_at").order("created_at");
      if (connError) return await fail("mailboxes/list", connError.message, 500, { userId: user.id });
      const { mailboxes, organization } = await withRoles(user.id, sb, data);
      return json({ mailboxes, connections, categories: CATEGORIES, organization });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action === "add_connection") return await addConnection(user.id, body);
      if (body.action === "update_connection") return await updateConnection(user.id, sb, body);
      if (body.action === "add_mailboxes") return await addMailboxes(user.id, sb, body);
      if (body.action === "find_shared_mailboxes") return await findSharedMailboxes(user.id, sb, body);
      if (body.action === "discover_shared") return await discoverShared(user.id, sb, body);
      if (body.action === "add_shared") return await addShared(user.id, body);
      if (body.action === "auto_shared") return await autoShared(user.id, sb, body);
      if (body.action === "set_mailbox_role") return await setMailboxRole(user.id, body);
      if (body.action === "send_forward_report") return await sendReportNow(user.id, body.mailbox_id);
      if (body.action === "send_weekly_report") return await sendWeeklyReportNow(user.id, body.mailbox_id);
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
      const hint = String(body.login_hint ?? "").trim().toLowerCase() || undefined;
      if (hint && !EMAIL.test(hint)) return json({ error: "login_hint must be a valid email address" }, 400);
      const s: State = { uid: user.id, days, origin, popup: body.popup === true, provider, shared, hint, auto: body.auto === true };
      return json({ url: OAUTH[provider].authorizeUrl(await signState(s), s) });
    }

    if (req.method === "PATCH") {
      const body = await req.json().catch(() => ({}));
      if (body.id && body.schedule) return await saveSchedule(user.id, body.id, body.schedule);
      if (body.id && body.forward_report) return await saveForwardReport(user.id, body.id, body.forward_report);
      if (body.id && body.weekly_report) return await saveWeeklyReport(user.id, body.id, body.weekly_report);
      if (!body.id) return json({ error: "id is required" }, 400);
      if (body.org_shared !== undefined) return await shareMailbox(user.id, body);
      return await saveSettings(user.id, body);
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
      const role = await mailboxRole(user.id, id);
      if (role && role !== "owner") return json({ error: "Only the person who connected this mailbox can disconnect it" }, 403);
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

// --- organization mailboxes ------------------------------------------------------------------------------
// See migrations/*_organization_mailboxes.sql. Rights come from mailbox_role(); writes use the service role.

type MailboxRole = "owner" | "admin" | "member";
// "none": no access (the mailbox is hidden from them, see migrations/*_mailbox_no_access_role.sql).
const MAILBOX_ROLES = ["admin", "member", "none"];
const canManage = (role: MailboxRole | null) => role === "owner" || role === "admin";
const ONLY_ADMINS = "Only admins of this mailbox can change it. Ask the person who connected it or an organization admin.";

// The user's role on mailbox `id`, or null without access.
async function mailboxRole(userId: string, id: unknown): Promise<MailboxRole | null> {
  const { data, error } = await adminClient().rpc("mailbox_role", { p_mailbox_id: id, p_user_id: userId });
  if (error) throw new Error(error.message);
  return data;
}

// Display names (profile name, else email) and emails of these users.
async function profilesOf(ids: string[]): Promise<Map<string, { email: string | null; full_name: string | null }>> {
  const unique = [...new Set(ids)];
  if (!unique.length) return new Map();
  const { data, error } = await adminClient().from("profiles").select("id, email, full_name").in("id", unique);
  if (error) throw new Error(error.message);
  return new Map(data.map((p) => [p.id, p]));
}

// Adds the caller's role and the owner's name to each mailbox (the same rules as mailbox_role(), without a query per
// mailbox; RLS has already left out mailboxes the caller can't see), and returns the caller's organization.
async function withRoles(userId: string, sb: SupabaseClient, rows: Record<string, unknown>[]) {
  const db = adminClient();
  const [own, set] = await Promise.all([
    db.from("organization_members").select("org_id, role, organizations(name)").eq("user_id", userId).maybeSingle(),
    sb.from("mailbox_members").select("mailbox_id, role").eq("user_id", userId),
  ]);
  for (const r of [own, set]) if (r.error) throw new Error(r.error.message);
  const roles = new Map((set.data ?? []).map((r) => [r.mailbox_id, r.role]));
  const owners = await profilesOf(rows.map((m) => m.owner_user_id as string).filter((id) => id !== userId));
  const mailboxes = rows.map((m) => {
    const owner = owners.get(m.owner_user_id as string);
    return {
      ...m,
      // Shared only with the caller's own organization; one left behind in an organization they're no longer in
      // (see *_organization_mailbox_membership.sql) shows as not shared.
      org_id: m.org_id && m.org_id === own.data?.org_id ? m.org_id : null,
      role: m.owner_user_id === userId ? "owner" : own.data?.role === "admin" ? "admin" : roles.get(m.id) ?? "member",
      owner_name: owner ? owner.full_name?.trim() || owner.email : null,
    };
  });
  const org = own.data as { org_id: string; role: string; organizations: { name: string } | null } | null;
  return { mailboxes, organization: org ? { id: org.org_id, name: org.organizations?.name ?? "", role: org.role } : null };
}

// Shares the caller's own mailbox with their organization, or stops sharing it. Roles set on it are dropped when it
// stops being shared. Shared mailboxes are always shared (one per organization) and personal ones never are, so in
// practice only stopping to share a personal mailbox shared before that rule changes anything.
async function shareMailbox(userId: string, body: Record<string, unknown>): Promise<Response> {
  if (typeof body.org_shared !== "boolean") return json({ error: "org_shared must be true or false" }, 400);
  const role = await mailboxRole(userId, body.id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (role !== "owner") return json({ error: "Only the person who connected this mailbox can change who it is shared with" }, 403);

  const db = adminClient();
  const { data: box, error: boxError } = await db.from("mailboxes").select("delegate_address, auth_type").eq("id", body.id).single();
  if (boxError) return await fail("mailboxes/share", boxError.message, 500, { userId, context: { id: body.id } });
  const sharedType = !!box.delegate_address || box.auth_type === "app";
  if (!body.org_shared && sharedType) {
    return json({ error: "Shared mailboxes are always shared with your organization" }, 400);
  }
  // Personal mailboxes stay private (enforced in the database too: mailboxes_only_shared_in_organization).
  if (body.org_shared && !sharedType) {
    return json({ error: "Only shared mailboxes can be shared with your organization. Personal mailboxes are visible only to you." }, 400);
  }
  let orgId: string | null = null;
  if (body.org_shared) {
    const { data: own, error } = await db.from("organization_members").select("org_id").eq("user_id", userId).maybeSingle();
    if (error) return await fail("mailboxes/share", error.message, 500, { userId });
    if (!own) return json({ error: "You're not in an organization" }, 400);
    orgId = own.org_id;
  }
  const { data, error } = await db.from("mailboxes").update({ org_id: orgId, updated_at: new Date().toISOString() })
    .eq("id", body.id).select("id, org_id, email_address").single();
  if (error && /mailboxes_org_address_key/.test(error.message)) return json({ error: "Your organization already has this mailbox" }, 400);
  if (error) return await fail("mailboxes/share", error.message, 500, { userId, context: { id: body.id } });
  if (!orgId) await db.from("mailbox_members").delete().eq("mailbox_id", data.id);
  await logApi("mailboxes/share", `${data.email_address} ${orgId ? "shared with organization" : "no longer shared"}`, {
    userId, level: "info", context: { mailbox: data.email_address, org_id: orgId },
  });
  return json({ mailbox: { id: data.id, org_id: data.org_id } });
}

// Everyone in an organization mailbox's organization with their role on it. `fixed`: the owner and organization
// admins, who are always admins of it.
async function mailboxAccess(userId: string, id: string | null): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  const db = adminClient();
  const { data: mailbox, error } = await db.from("mailboxes").select("id, owner_user_id, org_id").eq("id", id).single();
  if (error) return await fail("mailboxes/access", error.message, 500, { userId, context: { id } });
  if (!mailbox.org_id) return json({ error: "This mailbox isn't shared with an organization" }, 400);
  const [members, set] = await Promise.all([
    db.from("organization_members").select("user_id, role").eq("org_id", mailbox.org_id),
    db.from("mailbox_members").select("user_id, role").eq("mailbox_id", mailbox.id),
  ]);
  for (const r of [members, set]) if (r.error) return await fail("mailboxes/access", r.error.message, 500, { userId, context: { id } });
  const roles = new Map(set.data!.map((r) => [r.user_id, r.role]));
  const profiles = await profilesOf(members.data!.map((m) => m.user_id));
  const people = members.data!.map((m) => {
    const owner = m.user_id === mailbox.owner_user_id;
    return {
      user_id: m.user_id,
      email: profiles.get(m.user_id)?.email ?? null,
      full_name: profiles.get(m.user_id)?.full_name ?? null,
      role: owner ? "owner" : m.role === "admin" ? "admin" : roles.get(m.user_id) ?? "member",
      fixed: owner || m.role === "admin",
      org_admin: m.role === "admin",
      you: m.user_id === userId,
    };
  });
  const rank = (p: { role: string }) => ["owner", "admin", "member", "none"].indexOf(p.role);
  people.sort((a, b) => rank(a) - rank(b) || (a.full_name || a.email || "").localeCompare(b.full_name || b.email || ""));
  return json({ people, manage: canManage(role) });
}

async function setMailboxRole(userId: string, body: Record<string, unknown>): Promise<Response> {
  const role = String(body.role ?? "");
  if (!MAILBOX_ROLES.includes(role)) return json({ error: "role must be admin, member or none" }, 400);
  const mine = await mailboxRole(userId, body.mailbox_id);
  if (!mine) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(mine)) return json({ error: ONLY_ADMINS }, 403);
  const targetId = String(body.user_id ?? "");
  const db = adminClient();
  // By organization, not mailbox_role: someone with no access has no role but can be given one back.
  const [{ data: box, error: boxError }, { data: member, error: memberError }] = await Promise.all([
    db.from("mailboxes").select("owner_user_id, org_id").eq("id", body.mailbox_id).single(),
    db.from("organization_members").select("org_id, role").eq("user_id", targetId).maybeSingle(),
  ]);
  if (boxError || memberError) return await fail("mailboxes/set_role", (boxError ?? memberError)!.message, 500, { userId });
  if (!box.org_id || member?.org_id !== box.org_id) return json({ error: "That person isn't in this mailbox's organization" }, 404);
  if (box.owner_user_id === targetId) return json({ error: "The person who connected the mailbox is always its admin" }, 400);
  if (member?.role === "admin") return json({ error: "Organization admins are always admins of its mailboxes" }, 400);
  const { error } = await db.from("mailbox_members").upsert(
    { mailbox_id: body.mailbox_id, user_id: targetId, role, updated_by: userId, updated_at: new Date().toISOString() },
    { onConflict: "mailbox_id,user_id" },
  );
  if (error) return await fail("mailboxes/set_role", error.message, 500, { userId, context: { mailbox_id: body.mailbox_id } });
  return json({ member: { user_id: targetId, role } });
}

// Recipients typed as one or more addresses (commas, semicolons or spaces between): lowercase, de-duplicated, checked.
function parseAddresses(value: unknown, label: string): string[] | Response {
  const addresses = [...new Set(String(value ?? "").split(/[\s,;]+/).map((a) => a.trim().toLowerCase()).filter(Boolean))];
  const bad = addresses.find((a) => !EMAIL.test(a));
  if (bad) return json({ error: `${label}: ${bad} is not a valid email address` }, 400);
  if (addresses.length > MAX_SUMMARY_RECIPIENTS) return json({ error: `${label}: at most ${MAX_SUMMARY_RECIPIENTS} addresses` }, 400);
  return addresses;
}

// The daily forwarded-invoices report of a mailbox (owner or mailbox admins): on or off, the local time it goes out
// and to whom. Next / last times are written here and by the scheduler only.
async function saveForwardReport(userId: string, id: string, r: Record<string, unknown>): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);
  const enabled = r.enabled === true;
  const time = String(r.time ?? "");
  const timezone = String(r.timezone ?? "");
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) return json({ error: "time must be HH:MM" }, 400);
  if (!isValidTimeZone(timezone)) return json({ error: "Unknown time zone" }, 400);
  const addresses = parseAddresses(r.addresses, "Report recipients");
  if (addresses instanceof Response) return addresses;
  if (enabled && !addresses.length) return json({ error: "Enter at least one email address to send the report to" }, 400);

  const { data, error } = await adminClient().from("mailboxes").update({
    forward_report_enabled: enabled,
    forward_report_time: time,
    forward_report_timezone: timezone,
    forward_report_addresses: addresses.join(", ") || null,
    forward_report_next_at: enabled ? nextReportAt(time, timezone, new Date()).toISOString() : null,
    updated_at: new Date().toISOString(),
  }).eq("id", id)
    .select("id, forward_report_enabled, forward_report_time, forward_report_timezone, forward_report_addresses, forward_report_next_at, forward_report_last_at, forward_report_last_result")
    .single();
  if (error) return await fail("mailboxes/forward_report", error.message, 500, { userId, context: { id } });
  return json({ mailbox: data });
}

// Sends the forwarded-invoices report now (owner or mailbox admins), covering the time since the last one.
async function sendReportNow(userId: string, id: unknown): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);
  const db = adminClient();
  const { data, error } = await db.from("mailboxes").select(REPORT_MAILBOX_COLUMNS).eq("id", id).single();
  if (error) return await fail("mailboxes/forward_report", error.message, 500, { userId, context: { id } });
  const m = data as unknown as ReportMailbox & { forward_report_next_at: string | null };
  const me = (await profilesOf([userId])).get(userId);
  const result = await sendForwardReport(db, m, { id: userId, name: me?.full_name?.trim() || me?.email || null });
  // Sending early doesn't move the daily time; a report that is off stays off.
  const { data: box } = await db.from("mailboxes").select("forward_report_enabled").eq("id", m.id).single();
  if (!box?.forward_report_enabled) await db.from("mailboxes").update({ forward_report_next_at: null }).eq("id", m.id);
  return json({ result });
}

// --- Weekly forwarded-invoices report (_shared/weekly-report.ts), separate from the daily one above. ---

const WEEKLY_REPORT_COLUMNS = "id, weekly_report_enabled, weekly_report_day, weekly_report_time, weekly_report_timezone, " +
  "weekly_report_addresses, weekly_report_next_at, weekly_report_last_at, weekly_report_last_result";

// Adds the weekly report settings to the listed mailboxes. Read on their own, so the list still loads (without them)
// if the weekly report's columns aren't there yet.
async function addWeeklyReportSettings(userId: string, sb: SupabaseClient, mailboxes: Record<string, unknown>[] | null) {
  if (!mailboxes?.length) return;
  const { data, error } = await sb.from("mailboxes").select(WEEKLY_REPORT_COLUMNS).in("id", mailboxes.map((m) => m.id));
  if (error) return await logApi("mailboxes/weekly_report", error.message, { userId, level: "warn" });
  const byId = new Map((data as unknown as Record<string, unknown>[]).map((w) => [w.id, w]));
  for (const m of mailboxes) Object.assign(m, byId.get(m.id) ?? {});
}

// The weekly report of a mailbox (owner or mailbox admins): on or off, the day and local time it goes out, and to whom.
async function saveWeeklyReport(userId: string, id: string, r: Record<string, unknown>): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);
  const enabled = r.enabled === true;
  const day = Number(r.day);
  const time = String(r.time ?? "");
  const timezone = String(r.timezone ?? "");
  if (!Number.isInteger(day) || day < 0 || day > 6) return json({ error: "day must be 0 (Sunday) to 6 (Saturday)" }, 400);
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) return json({ error: "time must be HH:MM" }, 400);
  if (!isValidTimeZone(timezone)) return json({ error: "Unknown time zone" }, 400);
  const addresses = parseAddresses(r.addresses, "Weekly report recipients");
  if (addresses instanceof Response) return addresses;
  if (enabled && !addresses.length) return json({ error: "Enter at least one email address to send the weekly report to" }, 400);

  const { data, error } = await adminClient().from("mailboxes").update({
    weekly_report_enabled: enabled,
    weekly_report_day: day,
    weekly_report_time: time,
    weekly_report_timezone: timezone,
    weekly_report_addresses: addresses.join(", ") || null,
    weekly_report_next_at: enabled ? nextWeeklyReportAt(day, time, timezone, new Date()).toISOString() : null,
    updated_at: new Date().toISOString(),
  }).eq("id", id).select(WEEKLY_REPORT_COLUMNS).single();
  if (error) return await fail("mailboxes/weekly_report", error.message, 500, { userId, context: { id } });
  return json({ mailbox: data, day_name: dayName(day) });
}

// Sends the weekly report now (owner or mailbox admins), for the past 7 days; the weekly day and time don't move.
async function sendWeeklyReportNow(userId: string, id: unknown): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);
  const db = adminClient();
  const { data, error } = await db.from("mailboxes").select(WEEKLY_REPORT_MAILBOX_COLUMNS).eq("id", id).single();
  if (error) return await fail("mailboxes/weekly_report", error.message, 500, { userId, context: { id } });
  const me = (await profilesOf([userId])).get(userId);
  const result = await sendWeeklyReport(db, data as unknown as WeeklyReportMailbox, { id: userId, name: me?.full_name?.trim() || me?.email || null });
  return json({ result });
}

// Saves a mailbox's processing schedule and computes its next run. Schedule columns are written with the
// service role (after checking the caller is the owner or a mailbox admin) so users can't set next/last run times
// directly.
async function saveSchedule(userId: string, id: string, s: Record<string, unknown>): Promise<Response> {
  const role = await mailboxRole(userId, id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);

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

    const id = await saveMailbox(state.uid, {
      provider: state.provider, auth_type: "oauth", connection_id: null, email_address: email, delegate_address: delegate,
      display_name: displayName,
    }, refreshToken, new Date(Date.now() + state.days * 86_400_000).toISOString());
    // An own Microsoft mailbox: the dashboard then offers the shared mailboxes this account can open (or, after
    // signing in to the app with Microsoft, adds them all).
    const own = state.provider === "microsoft" && !delegate ? { mailbox_id: id, ...(state.auto ? { auto: "1" } : {}) } : {};
    return back({ connected: email, days: String(state.days), ...own });
  } catch (err) {
    if (err instanceof ConnectError) return await failed(err.message, err.detail, err.context);
    const message = err instanceof Error ? err.message : String(err);
    return await failed(`Could not save the mailbox: ${message}`, err);
  }
}

// Creates or updates the user's signed-in mailbox row and stores its refresh token in Vault until `expiresAt`.
async function saveMailbox(
  uid: string,
  fields: {
    provider: Provider; auth_type: "oauth"; connection_id: null; email_address: string; delegate_address: string | null;
    display_name: string | null;
  },
  secret: string,
  expiresAt: string,
): Promise<string> {
  const db = adminClient();
  const { data: existing, error: findError } = await db.from("mailboxes").select("id, credentials_secret_id")
    .eq("owner_user_id", uid).eq("email_address", fields.email_address).maybeSingle();
  if (findError) throw new Error(findError.message);

  let id = existing?.id as string | undefined;
  if (!id) {
    const shared = await settingsFromOtherUsers(db, uid, fields.email_address);
    // A shared mailbox (read through the user's account): one copy in the whole app, shared with the organization.
    const org = fields.delegate_address ? await organizationCopy(db, uid, fields.email_address) : null;
    if (org?.refused) throw new ConnectError(org.refused);
    const { data, error } = await db.from("mailboxes")
      .insert({ owner_user_id: uid, ...fields, ...shared, ...(org?.orgId ? { org_id: org.orgId } : {}) }).select("id").single();
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
    access_expires_at: expiresAt,
    updated_at: new Date().toISOString(),
  }).eq("id", id);
  if (error) throw new Error(error.message);
  return id;
}

// For a shared mailbox the user is adding: their organization (null if none), and why it can't be added (null if it
// can). A shared mailbox exists once in the whole app: the first copy added is the one everyone uses (everyone in its
// owner's organization; see migrations/*_one_shared_mailbox_per_organization.sql), so any other copy refuses the add.
// Copies found by address whoever owns them and whether or not they are marked shared yet.
async function organizationCopy(db: SupabaseClient, uid: string, address: string): Promise<{ orgId: string | null; refused: string | null }> {
  const { data: own, error } = await db.from("organization_members").select("org_id").eq("user_id", uid).maybeSingle();
  if (error) throw new Error(error.message);
  const orgId: string | null = own?.org_id ?? null;
  const { data: copies, error: copyError } = await db.from("mailboxes").select("owner_user_id")
    .eq("email_address", address).neq("owner_user_id", uid)
    .or("delegate_address.not.is.null,auth_type.eq.app,org_id.not.is.null").order("created_at").limit(50);
  if (copyError) throw new Error(copyError.message);
  if (!copies?.length) return { orgId, refused: null };

  const owners = [...new Set(copies.map((c) => c.owner_user_id))];
  const { data: colleagues, error: membersError } = orgId
    ? await db.from("organization_members").select("user_id").eq("org_id", orgId).in("user_id", owners)
    : { data: [], error: null };
  if (membersError) throw new Error(membersError.message);
  const colleague = owners.find((id) => (colleagues ?? []).some((c) => c.user_id === id));
  if (!colleague) {
    return {
      orgId,
      refused: `${address} is already connected by someone else. To use it, ask them to add you to their organization.`,
    };
  }
  const owner = (await profilesOf([colleague])).get(colleague);
  const by = owner?.full_name?.trim() || owner?.email || "a colleague";
  return { orgId, refused: `${address} is already in your organization (added by ${by}). Find it under Organization mailboxes.` };
}

// Settings of the same mailbox as another member of the user's organization most recently configured it, for a
// mailbox this user adds for the first time (typically a shared AP/AR mailbox colleagues already set up). Empty when
// the user is in no organization or no other member has it.
async function settingsFromOtherUsers(db: SupabaseClient, uid: string, address: string): Promise<Record<string, unknown>> {
  const { data: own, error: orgError } = await db.from("organization_members").select("org_id").eq("user_id", uid).maybeSingle();
  if (orgError) throw new Error(orgError.message);
  if (!own) return {};
  const { data: members, error: membersError } = await db.from("organization_members").select("user_id")
    .eq("org_id", own.org_id).neq("user_id", uid);
  if (membersError) throw new Error(membersError.message);
  if (!members.length) return {};
  const { data, error } = await db.from("mailboxes")
    .select("mailbox_type, forward_categories, smart_ap_address, send_summary, summary_address")
    .eq("email_address", address).in("owner_user_id", members.map((m) => m.user_id))
    .order("updated_at", { ascending: false }).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data ?? {};
}

// Per-mailbox settings (the owner or a mailbox admin). forward_categories must be categories of the mailbox's type;
// changing the type without sending forward_categories resets them to that type's default.
async function saveSettings(userId: string, body: Record<string, unknown>): Promise<Response> {
  const role = await mailboxRole(userId, body.id);
  if (!role) return json({ error: "Mailbox not found" }, 404);
  if (!canManage(role)) return json({ error: ONLY_ADMINS }, 403);
  const db = adminClient();
  const { data: current, error: readError } = await db.from("mailboxes").select("id, mailbox_type").eq("id", body.id).maybeSingle();
  if (readError) return await fail("mailboxes/update", readError.message, 500, { userId, context: { id: body.id } });
  if (!current) return json({ error: "Mailbox not found" }, 404);

  const update: Record<string, unknown> = { updated_at: new Date().toISOString() };
  if (body.smart_ap_address !== undefined) {
    const address = String(body.smart_ap_address ?? "").trim().toLowerCase();
    if (address && !EMAIL.test(address)) return json({ error: "Destination must be a valid email address" }, 400);
    update.smart_ap_address = address || null;
  }
  if (body.send_summary !== undefined) {
    if (typeof body.send_summary !== "boolean") return json({ error: "send_summary must be true or false" }, 400);
    update.send_summary = body.send_summary;
  }
  // Summary recipients: one or more addresses (commas, semicolons or spaces between), stored comma-separated; ""
  // clears them (own mailbox: itself, shared mailbox: the owner's account email).
  if (body.summary_address !== undefined) {
    const addresses = parseAddresses(body.summary_address, "Summary email");
    if (addresses instanceof Response) return addresses;
    update.summary_address = addresses.join(", ") || null;
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

  const { data, error } = await db.from("mailboxes").update(update).eq("id", body.id)
    .select("id, mailbox_type, forward_categories, smart_ap_address, send_summary, summary_address").maybeSingle();
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
    // A new one: one copy in the whole app (see organizationCopy), shared with everyone in the organization.
    let saveError;
    if (existing) {
      ({ error: saveError } = await db.from("mailboxes").update(fields).eq("id", existing.id));
    } else {
      let org: { orgId: string | null; refused: string | null };
      try {
        org = await organizationCopy(db, userId, address);
      } catch (err) {
        return { address, ok: false, error: err instanceof Error ? err.message : String(err) };
      }
      if (org.refused) return { address, ok: false, error: org.refused };
      ({ error: saveError } = await db.from("mailboxes").insert({ owner_user_id: userId, email_address: address, ...fields, org_id: org.orgId }));
    }
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

// The caller's own signed-in Microsoft mailbox `id`, whose token reads the shared mailboxes; a Response on error.
async function signedInMailbox(userId: string, id: unknown): Promise<(MailboxRow & { access_expires_at: string | null }) | Response> {
  const { data, error } = await adminClient().from("mailboxes")
    .select("id, provider, auth_type, email_address, delegate_address, credentials_secret_id, connection_id, active, access_expires_at")
    .eq("id", id).eq("owner_user_id", userId).maybeSingle();
  if (error) return await fail("mailboxes/shared", error.message, 500, { userId, context: { id } });
  if (!data) return json({ error: "Mailbox not found" }, 404);
  if (data.provider !== "microsoft" || data.auth_type !== "oauth" || data.delegate_address) {
    return json({ error: "Shared mailboxes are found through an Azure mailbox you signed in to yourself" }, 400);
  }
  if (!data.active || !data.credentials_secret_id) return json({ error: `Access to ${data.email_address} has expired. Reconnect it first.` }, 400);
  return data;
}

// Shared mailboxes the account behind `mailbox_id` has Full Access to, without those the user already has.
// Not finding any is not an error: Autodiscover may be unavailable, and addresses can still be added by hand.
async function discoverShared(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const source = await signedInMailbox(userId, body.mailbox_id);
  if (source instanceof Response) return source;
  let found;
  try {
    found = await alternateMailboxes(adminClient(), source);
  } catch (err) {
    await logApi("mailboxes/discover_shared", err, { userId, level: "warn", context: { mailbox: source.email_address } });
    return json({
      mailboxes: [],
      warning: `Couldn't list the shared mailboxes of ${source.email_address} automatically (the Azure app may need the ` +
        "EWS.AccessAsUser.All permission). You can still add them by address.",
    });
  }
  // Left out: the user's own and organization's mailboxes, and shared mailboxes anyone else already has (one copy in
  // the whole app, see organizationCopy).
  const addresses = found.mailboxes.map((m) => m.address);
  const [mine, taken] = await Promise.all([
    sb.from("mailboxes").select("email_address"),
    addresses.length
      ? adminClient().from("mailboxes").select("email_address").in("email_address", addresses)
        .or("delegate_address.not.is.null,auth_type.eq.app,org_id.not.is.null")
      : Promise.resolve({ data: [], error: null }),
  ]);
  const error = mine.error ?? taken.error;
  if (error) return await fail("mailboxes/discover_shared", error.message, 500, { userId });
  const have = new Set([...(mine.data ?? []), ...(taken.data ?? [])].map((m) => m.email_address));
  const mailboxes = found.mailboxes.filter((m) => !have.has(m.address));
  if (mailboxes.length) return json({ mailboxes });

  // Nothing to offer: say why, and keep Exchange's reply in the logs.
  const already = found.mailboxes.length;
  await logApi("mailboxes/discover_shared", `Autodiscover listed ${already} shared mailbox(es) (${found.entries} alternative mailbox entries) for ${source.email_address}`, {
    userId, level: "info", context: { mailbox: source.email_address, already_added: already, response: found.response },
  });
  return json({
    mailboxes: [],
    warning: already
      ? `All ${already} shared mailbox${already === 1 ? "" : "es"} ${source.email_address} can open ${already === 1 ? "is" : "are"} already added.`
      : `Exchange didn't report any shared mailboxes for ${source.email_address}. It only lists mailboxes where Full Access ` +
        "was granted to this user directly with automapping on (not through a group, and not with -AutoMapping $false). " +
        "Add the shared mailbox by address below; access is checked when it's added.",
  });
}

// Adds shared mailboxes read through the signed-in account behind `mailbox_id`: each gets its own copy of that
// account's refresh token and the same access period, after checking the account can open its inbox.
async function addShared(userId: string, body: Record<string, unknown>): Promise<Response> {
  const source = await signedInMailbox(userId, body.mailbox_id);
  if (source instanceof Response) return source;
  const raw = Array.isArray(body.addresses) ? body.addresses.join("\n") : String(body.addresses ?? "");
  const addresses = [...new Set(raw.split(/[\s,;]+/).map((a) => a.trim().toLowerCase()).filter(Boolean))];
  if (!addresses.length) return json({ error: "Select at least one shared mailbox" }, 400);
  if (addresses.length > MAX_BULK) return json({ error: `Add at most ${MAX_BULK} mailboxes at a time` }, 400);
  const results = await addSharedThrough(userId, source, addresses);
  return results instanceof Response ? results : json({ results });
}

// After signing in to the app with Microsoft: adds every shared mailbox the account behind `mailbox_id` has Full
// Access to and the user doesn't have yet. Finding none is not an error (see discoverShared).
async function autoShared(userId: string, sb: SupabaseClient, body: Record<string, unknown>): Promise<Response> {
  const res = await discoverShared(userId, sb, body);
  if (!res.ok) return res;
  const found: { mailboxes: { address: string }[]; warning?: string } = await res.json();
  if (!found.mailboxes.length) return json({ results: [], warning: found.warning });
  const source = await signedInMailbox(userId, body.mailbox_id);
  if (source instanceof Response) return source;
  const results = await addSharedThrough(userId, source, found.mailboxes.slice(0, MAX_BULK).map((m) => m.address));
  return results instanceof Response ? results : json({ results });
}

type AddResult = { address: string; ok: boolean; error?: string };

async function addSharedThrough(
  userId: string, source: MailboxRow & { access_expires_at: string | null }, addresses: string[],
): Promise<AddResult[] | Response> {
  const db = adminClient();
  let token: string, refreshToken: string;
  try {
    token = await delegatedToken(db, source);
    refreshToken = await readRefreshToken(db, source);
  } catch (err) {
    return await fail("mailboxes/add_shared", err instanceof Error ? err.message : String(err), 400, { userId, context: { mailbox: source.email_address } });
  }
  const expiresAt = source.access_expires_at ?? new Date(Date.now() + DEFAULT_ACCESS_DAYS * 86_400_000).toISOString();

  const addOne = async (address: string): Promise<AddResult> => {
    if (!EMAIL.test(address)) return { address, ok: false, error: "Not a valid email address" };
    if (address === source.email_address) return { address, ok: false, error: "This is the signed-in mailbox itself" };
    const { data: existing, error: findError } = await db.from("mailboxes").select("connection_id")
      .eq("owner_user_id", userId).eq("email_address", address).maybeSingle();
    if (findError) return { address, ok: false, error: findError.message };
    if (existing?.connection_id) return { address, ok: false, error: "Already added through an organization" };

    const res = await openInbox(token, address);
    if (!res.ok) {
      const detail = (await res.text()).slice(0, 1000);
      await logApi("mailboxes/add_shared", `Graph ${res.status} for ${address}`, { userId, level: "warn", context: { via: source.email_address, response: detail } });
      return {
        address, ok: false,
        error: res.status === 404 || res.status === 403 || res.status === 401
          ? `${source.email_address} can't open this mailbox (needs Full Access; reconnect ${source.email_address} if it was connected before shared mailboxes were supported)`
          : `Graph error ${res.status}`,
      };
    }
    try {
      await saveMailbox(userId, {
        provider: "microsoft", auth_type: "oauth", connection_id: null, email_address: address,
        delegate_address: source.email_address, display_name: null,
      }, refreshToken, expiresAt);
      return { address, ok: true };
    } catch (err) {
      return { address, ok: false, error: err instanceof Error ? err.message : String(err) };
    }
  };

  const results: AddResult[] = [];
  const queue = [...addresses];
  await Promise.all(Array.from({ length: Math.min(CHECK_CONCURRENCY, queue.length) }, async () => {
    for (let a = queue.shift(); a; a = queue.shift()) results.push(await addOne(a));
  }));
  results.sort((a, b) => addresses.indexOf(a.address) - addresses.indexOf(b.address));
  return results;
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
      // Shared scopes on every sign-in, so shared mailboxes found afterwards can be read with the same token.
      scope: MS_SHARED_SCOPES,
      // With a hint (the account already signed in to the app), sign in as it without the account picker.
      ...(s.hint ? { login_hint: s.hint } : { prompt: "select_account" }),
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
          scope: MS_SHARED_SCOPES,
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
        throw new ConnectError(`Allow ${APP_NAME} to read and send Gmail on Google's consent screen, then try again.`,
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
  // Account to sign in with (login_hint); used only to build the sign-in URL.
  hint?: string;
  // Started automatically after signing in to the app with Microsoft: shared mailboxes are then added without asking.
  auto?: boolean;
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
    provider: PROVIDERS.includes(s.provider) ? s.provider : "microsoft", shared: typeof s.shared === "string" ? s.shared : undefined, auto: s.auto === true };
}
