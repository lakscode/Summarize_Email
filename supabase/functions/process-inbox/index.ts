// Process new AP emails for one or all mailboxes and send each owner a summary report.
//
// Called by pg_cron (header x-cron-secret) for every active mailbox, or by a signed-in user
// from the dashboard ("Run now") for their own mailboxes. Without a mailbox_id it fans out one
// invocation per mailbox; each mailbox run happens in the background (202 returned at once).
// Large inboxes are processed in batches: the function re-invokes itself with the same run_id
// until the backlog is empty, then sends the report.
import { processEmail } from "../_shared/process.ts";
import { providerFor } from "../_shared/providers/index.ts";
import { buildReport } from "../_shared/report.ts";
import type { Mailbox, MailSummary } from "../_shared/types.ts";
import { adminClient, callerUserId, corsHeaders, env, json, mapLimit } from "../_shared/util.ts";

declare const EdgeRuntime: { waitUntil(p: Promise<unknown>): void } | undefined;

const BATCH = Number(Deno.env.get("MAX_EMAILS_PER_INVOCATION") ?? 12);
const CONCURRENCY = Number(Deno.env.get("CLASSIFY_CONCURRENCY") ?? 3);

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const body = (await req.json().catch(() => ({}))) as { mailbox_id?: string; run_id?: string };
  const db = adminClient();
  const isCron = req.headers.get("x-cron-secret") === env("CRON_SECRET");
  const userId = isCron ? null : await callerUserId(req);
  if (!isCron && !userId) return json({ error: "Unauthorized" }, 401);

  let q = db.from("mailboxes").select("*").eq("active", true);
  if (userId) q = q.eq("owner_user_id", userId);
  if (body.mailbox_id) q = q.eq("id", body.mailbox_id);
  const { data: mailboxes, error } = await q;
  if (error) return json({ error: error.message }, 500);

  const list = (mailboxes ?? []) as Mailbox[];
  if (!body.mailbox_id) {
    // Fan out: one invocation per mailbox, so each gets its own time budget.
    await Promise.all(list.map((mb) => invokeSelf({ mailbox_id: mb.id })));
    return json({ dispatched: list.map((mb) => mb.email_address) }, 202);
  }
  if (!list.length) return json({ error: "Mailbox not found or inactive" }, 404);

  const work = runMailbox(list[0], body.run_id ?? null);
  if (typeof EdgeRuntime !== "undefined") {
    EdgeRuntime.waitUntil(work);
    return json({ mailbox: list[0].email_address, started: true }, 202);
  }
  return json(await work);
});

/** Start a processing invocation and wait only until it has been accepted. */
function invokeSelf(payload: { mailbox_id: string; run_id?: string }): Promise<void> {
  return fetch(`${env("SUPABASE_URL")}/functions/v1/process-inbox`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-cron-secret": env("CRON_SECRET") },
    body: JSON.stringify(payload),
  }).then((r) => r.body?.cancel()).catch((e) => console.error("process-inbox dispatch failed:", e));
}

async function runMailbox(mailbox: Mailbox, existingRunId: string | null) {
  const db = adminClient();
  let runId = existingRunId;
  let startedAt = new Date();
  if (runId) {
    const { data } = await db.from("runs").select("started_at").eq("id", runId).single();
    if (data) startedAt = new Date(data.started_at);
  } else {
    const { data, error } = await db.from("runs").insert({ mailbox_id: mailbox.id }).select("id, started_at").single();
    if (error) return { mailbox: mailbox.email_address, error: error.message };
    runId = data.id as string;
    startedAt = new Date(data.started_at);
  }

  const provider = await providerFor(db, mailbox).catch((e) => e as Error);
  if (provider instanceof Error) return await failRun(runId, mailbox, provider);

  try {
    const since = new Date(startedAt.getTime() - mailbox.lookback_hours * 3600_000);
    const listed = await provider.listMessages(since);

    // Skip anything already stored or already tagged as processed.
    const known = new Set<string>();
    const ids = listed.map((m) => m.id);
    for (let i = 0; i < ids.length; i += 200) {
      const { data } = await db.from("emails").select("provider_message_id")
        .eq("mailbox_id", mailbox.id).in("provider_message_id", ids.slice(i, i + 200));
      for (const r of data ?? []) known.add(r.provider_message_id);
    }
    const fresh = listed.filter((m) => !known.has(m.id) && !m.labels.includes(mailbox.processed_label));

    // A thread takes the category of its latest message; earlier ones are recorded as skipped.
    const latest = new Map<string, MailSummary>();
    for (const m of fresh) {
      const key = m.threadId ?? m.id;
      const cur = latest.get(key);
      if (!cur || m.receivedAt > cur.receivedAt) latest.set(key, m);
    }
    const toProcess = [...latest.values()].sort((a, b) => a.receivedAt.getTime() - b.receivedAt.getTime());
    const superseded = fresh.filter((m) => latest.get(m.threadId ?? m.id)?.id !== m.id);
    if (superseded.length) {
      await db.from("emails").upsert(
        superseded.map((m) => ({
          mailbox_id: mailbox.id,
          run_id: runId,
          provider_message_id: m.id,
          thread_id: m.threadId,
          received_at: m.receivedAt.toISOString(),
          category: "other",
          status: "skipped",
          summary: "Earlier message in a thread - handled with the latest message.",
        })),
        { onConflict: "mailbox_id,provider_message_id", ignoreDuplicates: true },
      );
    }

    const batch = toProcess.slice(0, BATCH);
    await mapLimit(batch, CONCURRENCY, async (m) => {
      try {
        const msg = await provider.getMessage(m.id);
        await processEmail(db, provider, mailbox, runId!, msg);
      } catch (err) {
        console.error(`Processing ${m.id} failed:`, err);
        await db.from("emails").upsert({
          mailbox_id: mailbox.id,
          run_id: runId,
          provider_message_id: m.id,
          thread_id: m.threadId,
          received_at: m.receivedAt.toISOString(),
          category: "other",
          status: "action_needed",
          employee_action: "Open this email and handle it manually - automatic processing failed.",
          error: err instanceof Error ? err.message : String(err),
        }, { onConflict: "mailbox_id,provider_message_id" });
      }
    });

    await db.rpc("increment_run_seen", { p_run_id: runId, p_by: batch.length + superseded.length });

    if (toProcess.length > batch.length) {
      continueLater(mailbox.id, runId!);
      return { mailbox: mailbox.email_address, run_id: runId, status: "continuing", remaining: toProcess.length - batch.length };
    }

    const report = await buildReport(db, mailbox, runId!);
    const { data: run } = await db.from("runs").select("emails_seen").eq("id", runId).single();
    let sentTo: string | null = null;
    if ((run?.emails_seen ?? 0) > 0 || report.actionCount > 0) {
      sentTo = mailbox.report_recipient || mailbox.email_address;
      await provider.sendHtml(sentTo, report.subject, report.html);
    }
    await db.from("runs").update({
      status: "completed",
      finished_at: new Date().toISOString(),
      counts: report.counts,
      report_sent_to: sentTo,
    }).eq("id", runId);
    await db.from("mailboxes").update({ last_run_at: new Date().toISOString() }).eq("id", mailbox.id);
    return { mailbox: mailbox.email_address, run_id: runId, status: "completed", counts: report.counts, report_sent_to: sentTo };
  } catch (err) {
    return await failRun(runId!, mailbox, err as Error);
  } finally {
    await provider.close?.();
  }
}

async function failRun(runId: string, mailbox: Mailbox, err: Error) {
  console.error(`Run for ${mailbox.email_address} failed:`, err);
  await adminClient().from("runs").update({
    status: "failed",
    finished_at: new Date().toISOString(),
    error: err.message,
  }).eq("id", runId);
  return { mailbox: mailbox.email_address, run_id: runId, status: "failed", error: err.message };
}

function continueLater(mailboxId: string, runId: string) {
  const call = invokeSelf({ mailbox_id: mailboxId, run_id: runId });
  if (typeof EdgeRuntime !== "undefined") EdgeRuntime.waitUntil(call);
}
