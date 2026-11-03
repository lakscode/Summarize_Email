// v1 copy of functions/scheduler (uses _shared_v1, UI in web/v1/). See migrations/*_organization_app_version.sql.
// Runs scheduled mailbox processing. Called every 5 minutes by pg_cron (see migrations *_schedules.sql, *_schedule_interval.sql) with the
// shared secret from Vault in the x-cron-secret header; not callable from the browser.
// For each mailbox whose schedule is due it processes every email not processed yet that arrived since the mailbox
// was connected (at most the last 14 days; older mail predates the app and is left alone). Batches of one mailbox
// run back to back while time allows; if mail is left over, the mailbox stays due and the next tick continues the
// same run, so the summary email goes out once, when the run is finished.
import { json } from "../_shared_v1/cors.ts";
import { logApi } from "../_shared_v1/log.ts";
import { processMailbox } from "../_shared_v1/process.ts";
import { type Frequency, type IntervalUnit, nextRun } from "../_shared_v1/schedule.ts";
import { adminClient } from "../_shared_v1/supabase.ts";

// Stop starting new mailboxes after this long so the invocation finishes inside the Edge Function limit.
const TIME_BUDGET_MS = 100_000;
// Mailboxes processed at once (each also classifies several emails at once, see process.ts), so dozens of AP/AR
// mailboxes on the same schedule are worked through in a few ticks.
const MAILBOX_CONCURRENCY = 3;
const MAX_LOOKBACK_HOURS = 24 * 14;
// An unfinished run is continued only if its last tick was this recent; otherwise (e.g. the schedule was off for a
// while) a new run starts.
const CONTINUE_WITHIN_MS = 60 * 60_000;
// This scheduler's app version (see migrations/*_organization_app_version.sql).
const APP_VERSION = "v1";

// Link target for the summary email. Ignored if APP_URL points at Supabase rather than the hosted UI.
const appUrl = (() => {
  const v = Deno.env.get("APP_URL")?.replace(/\/+$/, "");
  return v && /^https?:\/\//.test(v) && !v.includes(".supabase.co") ? v : null;
})();

Deno.serve(async (req) => {
  const db = adminClient();
  const { data: secret, error: secretError } = await db.rpc("cron_secret");
  if (secretError || !secret) {
    await logApi("scheduler", `Could not read cron secret: ${secretError?.message ?? "missing"}`);
    return json({ error: "Scheduler not configured" }, 500);
  }
  if (req.headers.get("x-cron-secret") !== secret) return json({ error: "Forbidden" }, 403);

  const started = Date.now();
  await db.rpc("expire_mailbox_access");
  const { data: due, error } = await db.from("mailboxes")
    .select("id, owner_user_id, email_address, created_at, schedule_last_result, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_anchor, schedule_interval, schedule_interval_unit, schedule_last_run_at, schedule_next_run_at")
    .eq("schedule_enabled", true)
    // Own credentials (sign-in), or an organization connection's API keys.
    .or("credentials_secret_id.not.is.null,connection_id.not.is.null")
    .lte("schedule_next_run_at", new Date().toISOString())
    .order("schedule_next_run_at")
    .limit(20);
  if (error) {
    await logApi("scheduler", error.message);
    return json({ error: error.message }, 500);
  }

  // Only mailboxes whose owner's organization is on this version; the other version's scheduler runs the rest.
  const owners = [...new Set((due ?? []).map((m) => m.owner_user_id))];
  const { data: memberships, error: versionError } = owners.length
    ? await db.from("organization_members").select("user_id, organizations(app_version)").in("user_id", owners)
    : { data: [], error: null };
  if (versionError) {
    await logApi("scheduler", versionError.message);
    return json({ error: versionError.message }, 500);
  }
  const versionOf = new Map((memberships ?? []).map((r) =>
    [r.user_id, (r.organizations as unknown as { app_version: string } | null)?.app_version ?? "current"]));

  const results: Record<string, unknown>[] = [];
  const queue = (due ?? []).filter((m) => (versionOf.get(m.owner_user_id) ?? "current") === APP_VERSION);
  await Promise.all(Array.from({ length: Math.min(MAILBOX_CONCURRENCY, queue.length) }, async () => {
    for (let m = queue.shift(); m; m = queue.shift()) {
      if (Date.now() - started > TIME_BUDGET_MS) return;
      const since = new Date(m.created_at);
      const hours = Math.min(Math.max(Math.ceil((Date.now() - since.getTime()) / 3600_000) + 1, 24), MAX_LOOKBACK_HOURS);
      const now = new Date();
      const following = nextRun({
        frequency: m.schedule_frequency as Frequency,
        days: m.schedule_days ?? [],
        time: String(m.schedule_time).slice(0, 5),
        timezone: m.schedule_timezone,
        anchor: m.schedule_anchor,
        interval: m.schedule_interval,
        intervalUnit: m.schedule_interval_unit as IntervalUnit,
      }, now);

      try {
        // A run the previous tick left unfinished continues (same started_at, so one summary email at the end).
        const previous = m.schedule_last_result as { at?: string; remaining?: number; started_at?: string } | null;
        const recent = !!previous?.at && Date.now() - new Date(previous.at).getTime() < CONTINUE_WITHIN_MS;
        let startedAt = recent && previous?.remaining && previous.started_at ? previous.started_at : null;
        let result: Awaited<ReturnType<typeof processMailbox>>;
        let processed = 0, batchMs = 0;
        const counts: Record<string, number> = {};
        do {
          const batchStart = Date.now();
          result = await processMailbox(m.owner_user_id, m.id, hours, false, "schedule", appUrl, { startedAt });
          batchMs = Date.now() - batchStart;
          startedAt = result.started_at;
          processed += result.processed;
          for (const [k, n] of Object.entries(result.counts)) counts[k] = (counts[k] ?? 0) + n;
          // Another batch only if it should finish inside the budget (judged by how long this one took), and not
          // when nothing could be taken (the rest is held by another run).
        } while (result.remaining > 0 && result.processed > 0 && Date.now() - started + batchMs < TIME_BUDGET_MS);

        const finished = result.remaining === 0;
        await db.from("mailboxes").update({
          // Unfinished runs stay due (and last_run_at unchanged) so the next tick continues.
          ...(finished ? { schedule_last_run_at: now.toISOString(), schedule_next_run_at: following.toISOString() } : {}),
          schedule_last_result: { at: now.toISOString(), ...result, processed, counts },
        }).eq("id", m.id);
        results.push({ ...result, processed, counts });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        await logApi("scheduler/run", err, { userId: m.owner_user_id, context: { mailbox: m.email_address } });
        // Skip to the next scheduled time rather than retrying a failing mailbox on every tick.
        await db.from("mailboxes").update({
          schedule_next_run_at: following.toISOString(),
          schedule_last_result: { at: now.toISOString(), error: message },
        }).eq("id", m.id);
        results.push({ mailbox: m.email_address, error: message });
      }
    }
  }));

  return json({ ran: results.length, due: due?.length ?? 0, results });
});
