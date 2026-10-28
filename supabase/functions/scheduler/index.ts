// Runs scheduled mailbox processing. Called every 5 minutes by pg_cron (see migrations *_schedules.sql, *_schedule_interval.sql) with the
// shared secret from Vault in the x-cron-secret header; not callable from the browser.
// For each mailbox whose schedule is due it processes mail received since the previous scheduled run. If a
// mailbox has more mail than one batch, it stays due and the next tick continues where this one stopped.
import { json } from "../_shared/cors.ts";
import { logApi } from "../_shared/log.ts";
import { processMailbox } from "../_shared/process.ts";
import { type Frequency, type IntervalUnit, nextRun } from "../_shared/schedule.ts";
import { adminClient } from "../_shared/supabase.ts";

// Stop starting new mailboxes after this long so the invocation finishes inside the Edge Function limit.
const TIME_BUDGET_MS = 100_000;
// Mailboxes processed at once (each also classifies several emails at once, see process.ts), so dozens of AP/AR
// mailboxes on the same schedule are worked through in a few ticks.
const MAILBOX_CONCURRENCY = 3;
const MAX_LOOKBACK_HOURS = 24 * 14;

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
    .select("id, owner_user_id, email_address, schedule_frequency, schedule_days, schedule_time, schedule_timezone, schedule_anchor, schedule_interval, schedule_interval_unit, schedule_last_run_at, schedule_next_run_at")
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

  const results: Record<string, unknown>[] = [];
  const queue = [...(due ?? [])];
  await Promise.all(Array.from({ length: Math.min(MAILBOX_CONCURRENCY, queue.length) }, async () => {
    for (let m = queue.shift(); m; m = queue.shift()) {
      if (Date.now() - started > TIME_BUDGET_MS) return;
      const since = m.schedule_last_run_at ? new Date(m.schedule_last_run_at) : new Date(Date.now() - 24 * 3600_000);
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
        const result = await processMailbox(m.owner_user_id, m.id, hours, false, "schedule", appUrl);
        const finished = result.remaining === 0;
        await db.from("mailboxes").update({
          // Unfinished batches keep the run due (and last_run_at unchanged) so the next tick continues.
          ...(finished ? { schedule_last_run_at: now.toISOString(), schedule_next_run_at: following.toISOString() } : {}),
          schedule_last_result: { at: now.toISOString(), ...result },
        }).eq("id", m.id);
        results.push(result);
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
