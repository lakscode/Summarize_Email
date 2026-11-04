// AP inbox processing for the signed-in user's connected mailboxes and the organization mailboxes they have a role on
// (any role may process and mark emails; see migrations/*_organization_mailboxes.sql).
//   GET   ?status=&mailbox_id=                          -> { emails: [...] } newest first
//   GET   ?preview=<email id>                           -> { email: {...} } the message read live from the mailbox
//   GET   ?history=<email id>                           -> { history: [...] } oldest first: each processing run and
//         each time someone marked it done / reopened it, with who and when (email_processing_log)
//   POST  {action: "process", mailbox_id, lookback_hours, reprocess?, all?, started_at?} -> process a batch now
//         (see _shared/process.ts); with all, the page calls again until remaining is 0
//   POST  {action: "preview", mailbox_id, lookback_hours} -> { new, already_processed, ... } to confirm a re-check first
//   PATCH {id, status: "done" | "action_needed"}       -> employee marks an item handled / reopens it; 409 when it
//         already has that status (someone else changed it first), naming who
//   PATCH {id, category}                                -> corrects the category (see changeCategory)
//   GET   ?assignees=<mailbox id>                       -> { people: [{id, name, email}] } who an email can be assigned to
//   GET   ?assigned=me&mailbox_id=                      -> { emails, unseen } the caller's open tasks (not done), in that
//                                                          mailbox (the Inbox tab) or, without mailbox_id, any (the bell)
//   PATCH {id, assign_to: <user id> | null}             -> assigns the email to someone / unassigns it (see assignEmail)
//   POST  {action: "seen", mailbox_id?}                 -> the caller has seen their assigned tasks (of that mailbox only,
//                                                          when given), so they no longer count as new on the bell
//   POST  {action: "check_statement", id}               -> checks a statement's invoices again (_shared/statement.ts), e.g.
//         after more invoices came in -> { statement_check, statement_checked_at }
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi, logStep } from "../_shared/log.ts";
import { previewEmail, previewMailbox, processMailbox, processorOf, STATUSES } from "../_shared/process.ts";
import { CATEGORIES, type MailboxType } from "../_shared/classify.ts";
import { isNotFound } from "../_shared/mail.ts";
import { CATEGORY_LABELS } from "../_shared/report.ts";
import { checkStatement } from "../_shared/statement.ts";
import { adminClient, callerUser } from "../_shared/supabase.ts";
import type { SupabaseClient, User } from "jsr:@supabase/supabase-js@2";

// The email's details copied into its email_processing_log row (a full snapshot, as process.ts writes); also in the
// PATCH select below.
const SNAPSHOT_COLUMNS =
  "message_id, received_at, from_name, from_address, web_link, attachment_names, category, confidence, vendor, invoice_number, amount, due_date, summary, next_step";
// What the Inbox lists of each email.
const LIST_COLUMNS =
  "id, mailbox_id, received_at, from_name, from_address, subject, web_link, attachment_names, category, confidence, vendor, invoice_number, amount, due_date, summary, next_step, action_taken, status, error, processed_by_email, processed_by_name, processed_at, processed_trigger, category_changed_by_name, category_changed_at, assigned_to, assigned_to_name, assigned_by_name, assigned_at, assignment_seen_at, statement_invoices, statement_check, statement_checked_at";

// The people with access to the mailbox (mailbox_role: its owner, and the members of its organization who weren't given
// "no access"), who its emails can be assigned to.
async function mailboxPeople(mailboxId: string): Promise<{ id: string; name: string; email: string | null }[]> {
  const db = adminClient();
  const { data: mailbox, error } = await db.from("mailboxes").select("owner_user_id, org_id").eq("id", mailboxId).maybeSingle();
  if (error) throw new Error(error.message);
  if (!mailbox) return [];
  let ids = [mailbox.owner_user_id as string];
  if (mailbox.org_id) {
    // mailbox_role()'s rules in two queries (not one call per member): only while the owner is in the organization;
    // its admins always, other members unless given "no access".
    const [members, hidden] = await Promise.all([
      db.from("organization_members").select("user_id, role").eq("org_id", mailbox.org_id),
      db.from("mailbox_members").select("user_id").eq("mailbox_id", mailboxId).eq("role", "none"),
    ]);
    for (const r of [members, hidden]) if (r.error) throw new Error(r.error.message);
    const noAccess = new Set(hidden.data!.map((r) => r.user_id));
    if (members.data!.some((m) => m.user_id === mailbox.owner_user_id)) {
      ids = [...new Set([...ids, ...members.data!.filter((m) => m.role === "admin" || !noAccess.has(m.user_id)).map((m) => m.user_id as string)])];
    }
  }
  const { data: profiles, error: profilesError } = await db.from("profiles").select("id, email, full_name").in("id", ids);
  if (profilesError) throw new Error(profilesError.message);
  return profiles.map((p) => ({ id: p.id, name: p.full_name?.trim() || p.email || "Unknown user", email: p.email }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// Assigns the email to someone with access to its mailbox (null: unassigns it). The assignee is notified (their
// assignment_seen_at is cleared) unless they assigned it to themselves. Logged in email_processing_log like marking done.
async function assignEmail(user: User, sb: SupabaseClient, id: string, assignTo: string | null) {
  // The caller's own client first: RLS says whether they may see (and so assign) the email.
  const { data: email, error } = await sb.from("processed_emails")
    .select(`id, status, mailbox_id, owner_user_id, subject, assigned_to, assigned_to_name, mailboxes(email_address), ${SNAPSHOT_COLUMNS}`)
    .eq("id", id).maybeSingle();
  if (error) return await fail("inbox/assign", error.message, 500, { userId: user.id, context: { email_id: id } });
  if (!email) return json({ error: "Email not found" }, 404);
  if ((email.assigned_to ?? null) === assignTo) return json({ email: { id, assigned_to: assignTo } });
  let assignee: { id: string; name: string } | null = null;
  if (assignTo) {
    assignee = (await mailboxPeople(email.mailbox_id)).find((p) => p.id === assignTo) ?? null;
    if (!assignee) return json({ error: "That person has no access to this mailbox" }, 400);
  }

  const db = adminClient();
  const processor = await processorOf(db, user.id);
  const by = { processed_by: user.id, processed_by_email: processor.email ?? user.email ?? null, processed_by_name: processor.name ?? user.email ?? null };
  const at = new Date().toISOString();
  const { error: updateError } = await db.from("processed_emails").update({
    assigned_to: assignee?.id ?? null,
    assigned_to_name: assignee?.name ?? null,
    assigned_by: assignee ? user.id : null,
    assigned_by_name: assignee ? by.processed_by_name : null,
    assigned_at: assignee ? at : null,
    assignment_seen_at: assignee?.id === user.id ? at : null,
    updated_at: at,
  }).eq("id", id);
  if (updateError) return await fail("inbox/assign", updateError.message, 500, { userId: user.id, context: { email_id: id } });

  const action = assignee ? `Assigned to ${assignee.name}` : `Unassigned${email.assigned_to_name ? ` (was ${email.assigned_to_name})` : ""}`;
  const mailboxAddress = (email.mailboxes as { email_address?: string } | null)?.email_address ?? null;
  const logContext = { email_id: id, subject: email.subject, mailbox: mailboxAddress };
  const { error: logError } = await db.from("email_processing_log").insert({
    processed_email_id: id,
    mailbox_id: email.mailbox_id,
    mailbox_address: mailboxAddress,
    ...Object.fromEntries(SNAPSHOT_COLUMNS.split(", ").map((c) => [c, (email as Record<string, unknown>)[c]])),
    subject: email.subject,
    owner_user_id: email.owner_user_id,
    ...by,
    processed_at: at,
    trigger: "user",
    status: email.status,
    action,
  });
  if (logError) await logApi("inbox/assign", logError.message, { userId: user.id, level: "warn", context: logContext });
  await logStep("inbox/assign", `"${email.subject || "(no subject)"}": ${action.toLowerCase()} by ${by.processed_by_name ?? "unknown"}`,
    user.id, { ...logContext, assigned_to: assignee?.id ?? null });
  return json({ email: { id, assigned_to: assignee?.id ?? null, assigned_to_name: assignee?.name ?? null } });
}

// Someone corrects an email's category to another of its mailbox type's (classify.ts CATEGORIES). Unless it is done,
// its status follows: "other" has nothing to do, any other category needs action. Nothing is forwarded; a re-check
// run leaves the correction alone (process.ts recheckable). Logged in email_processing_log like marking done.
async function changeCategory(user: User, sb: SupabaseClient, id: string, category: string) {
  // The caller's own client first: RLS says whether they may see (and so change) the email.
  const { data: email, error } = await sb.from("processed_emails")
    .select(`id, status, mailbox_id, owner_user_id, subject, mailboxes(email_address, mailbox_type), ${SNAPSHOT_COLUMNS}`)
    .eq("id", id).maybeSingle();
  if (error) return await fail("inbox/category", error.message, 500, { userId: user.id, context: { email_id: id } });
  if (!email) return json({ error: "Email not found" }, 404);
  const mailbox = email.mailboxes as { email_address?: string; mailbox_type?: MailboxType } | null;
  const allowed = CATEGORIES[mailbox?.mailbox_type ?? "ap"];
  if (!allowed.includes(category)) return json({ error: `category must be one of: ${allowed.join(", ")}` }, 400);
  if (email.category === category) return json({ email: { id, category, status: email.status } });

  const db = adminClient();
  const processor = await processorOf(db, user.id);
  const by = { processed_by: user.id, processed_by_email: processor.email ?? user.email ?? null, processed_by_name: processor.name ?? user.email ?? null };
  const at = new Date().toISOString();
  const label = (c: string | null) => CATEGORY_LABELS[c ?? ""] ?? c ?? "none";
  const status = email.status === "done" ? "done" : category === "other" ? "no_action" : "action_needed";
  // The classifier's next step was for the old category; say what it is now when that changes what to do.
  const nextStep = status === email.status ? email.next_step
    : status === "no_action" ? `Changed to ${label(category)} by ${by.processed_by_name}: nothing to do.`
    : `Changed to ${label(category)} by ${by.processed_by_name}: check it and handle it.`;
  // Only from the category read above, so a change made meanwhile by someone else isn't overwritten unseen.
  const { data: updated, error: updateError } = await db.from("processed_emails").update({
    category, status, next_step: nextStep, ...by, processed_trigger: "user", processed_at: at, updated_at: at,
    category_changed_by: user.id, category_changed_by_name: by.processed_by_name, category_changed_at: at,
  }).eq("id", id).filter("category", email.category === null ? "is" : "eq", email.category).select("id");
  if (updateError) return await fail("inbox/category", updateError.message, 500, { userId: user.id, context: { email_id: id } });
  if (!updated?.length) return json({ error: "Someone else changed this email meanwhile. Reload and try again." }, 409);

  const action = `Category changed from ${label(email.category)} to ${label(category)}`;
  const logContext = { email_id: id, subject: email.subject, mailbox: mailbox?.email_address ?? null };
  const { error: logError } = await db.from("email_processing_log").insert({
    processed_email_id: id,
    mailbox_id: email.mailbox_id,
    mailbox_address: mailbox?.email_address ?? null,
    ...Object.fromEntries(SNAPSHOT_COLUMNS.split(", ").map((c) => [c, (email as Record<string, unknown>)[c]])),
    category,
    next_step: nextStep,
    subject: email.subject,
    owner_user_id: email.owner_user_id,
    ...by,
    processed_at: at,
    trigger: "user",
    status,
    action,
  });
  if (logError) await logApi("inbox/category", logError.message, { userId: user.id, level: "warn", context: logContext });
  await logStep("inbox/category", `"${email.subject || "(no subject)"}": ${action.toLowerCase()} by ${by.processed_by_name ?? "unknown"}`,
    user.id, { ...logContext, category, status });
  return json({ email: { id, category, status } });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user, sb } = caller;
  const url = new URL(req.url);

  try {
    if (req.method === "GET" && url.searchParams.has("preview")) {
      const id = url.searchParams.get("preview")!;
      try {
        const email = await previewEmail(user.id, id);
        return email ? json({ email }) : json({ error: "Email not found" }, 404);
      } catch (err) {
        // Gone from the inbox: deleted, or moved (Outlook gives a moved message a new id).
        if (isNotFound(err)) {
          return await fail("inbox/preview", "This email is no longer in the inbox. It may have been moved or deleted.", 404, { userId: user.id, context: { email_id: id } });
        }
        throw err;
      }
    }

    if (req.method === "GET" && url.searchParams.has("history")) {
      const id = url.searchParams.get("history")!;
      // The caller's own client: RLS limits it to emails of mailboxes they have a role on.
      const { data, error } = await sb.from("email_processing_log")
        .select("id, processed_at, processed_by_name, processed_by_email, trigger, status, action, error")
        .eq("processed_email_id", id)
        .order("processed_at", { ascending: true })
        .order("id", { ascending: true })
        .limit(200);
      if (error) return await fail("inbox/history", error.message, 500, { userId: user.id, context: { email_id: id } });
      return json({ history: data });
    }

    if (req.method === "GET" && url.searchParams.has("assignees")) {
      const mailboxId = url.searchParams.get("assignees")!;
      // Only for a mailbox the caller can see (RLS).
      const { data: mailbox } = await sb.from("mailboxes").select("id").eq("id", mailboxId).maybeSingle();
      if (!mailbox) return json({ error: "Mailbox not found" }, 404);
      return json({ people: await mailboxPeople(mailboxId) });
    }

    if (req.method === "GET" && url.searchParams.get("assigned") === "me") {
      // The caller's own client: an email of a mailbox they have since lost access to drops out.
      let query = sb.from("processed_emails").select(LIST_COLUMNS)
        .eq("assigned_to", user.id).neq("status", "done");
      const mailboxId = url.searchParams.get("mailbox_id");
      if (mailboxId) query = query.eq("mailbox_id", mailboxId);
      const { data, error } = await query.order("assigned_at", { ascending: false }).limit(300);
      if (error) return await fail("inbox/assigned", error.message, 500, { userId: user.id });
      return json({ emails: data, unseen: data.filter((e) => !e.assignment_seen_at).length });
    }

    if (req.method === "GET") {
      let query = sb
        .from("processed_emails")
        .select(LIST_COLUMNS)
        .order("received_at", { ascending: false })
        .limit(300);
      const status = url.searchParams.get("status");
      if (status && STATUSES.includes(status)) query = query.eq("status", status);
      const mailboxId = url.searchParams.get("mailbox_id");
      if (mailboxId) query = query.eq("mailbox_id", mailboxId);
      const { data, error } = await query;
      if (error) return await fail("inbox/list", error.message, 500, { userId: user.id });
      return json({ emails: data });
    }

    if (req.method === "PATCH") {
      const body = await req.json().catch(() => ({}));
      if (body.id && typeof body.category === "string") return await changeCategory(user, sb, body.id, body.category);
      if (body.id && "assign_to" in body) {
        if (body.assign_to !== null && typeof body.assign_to !== "string") return json({ error: "assign_to must be a user id or null" }, 400);
        return await assignEmail(user, sb, body.id, body.assign_to);
      }
      if (!body.id || !["done", "action_needed"].includes(body.status)) {
        return json({ error: "id and status (done | action_needed) are required" }, 400);
      }
      // Only from another status, so two people changing the same email at once can't both log the change.
      const { data, error } = await sb.from("processed_emails")
        .update({ status: body.status, updated_at: new Date().toISOString() })
        .eq("id", body.id)
        .neq("status", body.status)
        .select("id, status, mailbox_id, owner_user_id, subject, mailboxes(email_address), message_id, received_at, from_name, from_address, web_link, attachment_names, category, confidence, vendor, invoice_number, amount, due_date, summary, next_step")
        .maybeSingle();
      if (error) return await fail("inbox/update", error.message, 500, { userId: user.id, context: { email_id: body.id } });
      if (!data) {
        const { data: current } = await sb.from("processed_emails")
          .select("status, processed_by_name, processed_by_email").eq("id", body.id).maybeSingle();
        if (!current) return json({ error: "Email not found" }, 404);
        const who = current.processed_by_name || current.processed_by_email || "someone";
        return json({ error: `Already ${body.status === "done" ? "marked done" : "reopened"} by ${who}.` }, 409);
      }
      // Marking done / reopening counts as processing it: this user and time become the email's "last processed".
      const db = adminClient();
      const processor = await processorOf(db, user.id);
      const by = { processed_by: user.id, processed_by_email: processor.email ?? user.email ?? null, processed_by_name: processor.name ?? user.email ?? null };
      const processedAt = new Date().toISOString();
      const { error: byError } = await db.from("processed_emails").update({ ...by, processed_trigger: "user", processed_at: processedAt }).eq("id", data.id);
      const mailboxAddress = (data.mailboxes as { email_address?: string } | null)?.email_address ?? null;
      const logContext = { email_id: data.id, subject: data.subject, mailbox: mailboxAddress };
      if (byError) await logApi("inbox/update", byError.message, { userId: user.id, level: "warn", context: logContext });
      const { error: logError } = await db.from("email_processing_log").insert({
        processed_email_id: data.id,
        mailbox_id: data.mailbox_id,
        mailbox_address: mailboxAddress,
        ...Object.fromEntries(SNAPSHOT_COLUMNS.split(", ").map((c) => [c, (data as Record<string, unknown>)[c]])),
        subject: data.subject,
        owner_user_id: data.owner_user_id,
        ...by,
        processed_at: processedAt,
        trigger: "user",
        status: data.status,
        action: data.status === "done" ? "Marked done" : "Reopened",
      });
      if (logError) await logApi("inbox/update", logError.message, { userId: user.id, level: "warn", context: logContext });
      await logStep("inbox/update", `"${data.subject || "(no subject)"}" ${data.status === "done" ? "marked done" : "reopened"} by ${by.processed_by_name ?? "unknown"}`,
        user.id, { ...logContext, status: data.status });
      return json({ email: { id: data.id, status: data.status } });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action === "check_statement") {
        if (!body.id) return json({ error: "id is required" }, 400);
        // The caller's own client: RLS says whether they may see the statement.
        const { data: email, error } = await sb.from("processed_emails").select("id, mailbox_id, category, statement_invoices")
          .eq("id", body.id).maybeSingle();
        if (error) return await fail("inbox/statement", error.message, 500, { userId: user.id, context: { email_id: body.id } });
        if (!email) return json({ error: "Email not found" }, 404);
        if (email.category !== "statement" || !email.statement_invoices?.length) {
          return json({ error: "No invoice numbers were read from this statement. Process it again to read them." }, 400);
        }
        const db = adminClient();
        const check = await checkStatement(db, email.mailbox_id, email.id, email.statement_invoices);
        const at = new Date().toISOString();
        const { error: updateError } = await db.from("processed_emails").update({ statement_check: check, statement_checked_at: at }).eq("id", email.id);
        if (updateError) return await fail("inbox/statement", updateError.message, 500, { userId: user.id, context: { email_id: body.id } });
        return json({ statement_check: check, statement_checked_at: at });
      }
      if (body.action === "seen") {
        let seen = adminClient().from("processed_emails").update({ assignment_seen_at: new Date().toISOString() })
          .eq("assigned_to", user.id).is("assignment_seen_at", null);
        if (body.mailbox_id) seen = seen.eq("mailbox_id", body.mailbox_id);
        const { error } = await seen;
        if (error) return await fail("inbox/seen", error.message, 500, { userId: user.id });
        return json({ ok: true });
      }
      if (!["process", "preview"].includes(body.action) || !body.mailbox_id) {
        return json({ error: "action ('process' | 'preview') and mailbox_id are required" }, 400);
      }
      const hours = Math.min(Math.max(Number(body.lookback_hours) || 24, 1), 24 * 14);
      if (body.action === "preview") return json(await previewMailbox(user.id, body.mailbox_id, hours));
      const origin = req.headers.get("Origin");
      const appUrl = origin && /^https?:\/\/[^/]+$/.test(origin) ? origin : null;
      // all: the page keeps calling until nothing remains, passing back the first call's started_at.
      const startedAt = typeof body.started_at === "string" && !isNaN(Date.parse(body.started_at)) ? body.started_at : null;
      const untilDone = body.all === true ? { startedAt } : null;
      return json(await processMailbox(user.id, body.mailbox_id, hours, body.reprocess === true, "manual", appUrl, untilDone));
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("inbox", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
