// Dashboard actions on a single action item.
//   POST {email_id, action: "forward"}   send an invoice to Smart AP (the employee's confirmation)
//        {..., verified: true}           required to forward an item that was flagged as possible fraud
//   POST {email_id, action: "complete"}  mark done and tag the email as processed in the mailbox
//   POST {email_id, action: "reopen"}    put a completed item back on the action list
import { providerFor } from "../_shared/providers/index.ts";
import type { Mailbox } from "../_shared/types.ts";
import { adminClient, callerUserId, corsHeaders, json } from "../_shared/util.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });
  const userId = await callerUserId(req);
  if (!userId) return json({ error: "Unauthorized" }, 401);

  const { email_id, action, verified } = await req.json();
  const db = adminClient();
  const { data: item } = await db.from("emails").select("*, mailboxes(*)").eq("id", email_id).maybeSingle();
  const mailbox = item?.mailboxes as Mailbox | undefined;
  if (!item || !mailbox || mailbox.owner_user_id !== userId) return json({ error: "Not found" }, 404);

  const now = new Date().toISOString();
  try {
    if (action === "forward") {
      if (item.category !== "invoice") return json({ error: "Only invoices can be sent to Smart AP" }, 400);
      if (!mailbox.smart_ap_address) return json({ error: "Set the Smart AP address in mailbox settings first" }, 400);
      if (item.status === "flagged" && verified !== true) {
        return json({ error: "This invoice is flagged as possible fraud. Verify it with the vendor by phone first." }, 409);
      }
      if (item.forwarded_at) return json({ error: `Already sent to Smart AP on ${item.forwarded_at.slice(0, 10)}` }, 409);

      const provider = await providerFor(db, mailbox);
      try {
        const msg = await provider.getMessage(item.provider_message_id);
        await provider.forward(msg, mailbox.smart_ap_address,
          `Invoice ${item.invoice_number ?? ""} from ${item.vendor_name ?? msg.from.address} for processing.`);
        await provider.addLabel(item.provider_message_id, mailbox.processed_label).catch(() => {});
      } finally {
        await provider.close?.();
      }
      await db.from("emails").update({
        status: "done",
        forwarded_at: now,
        labeled_at: now,
        completed_at: now,
        completed_by: userId,
        action_taken: `Forwarded to Smart AP (${mailbox.smart_ap_address}) from the dashboard${item.status === "flagged" ? " after manual verification" : ""}.`,
        employee_action: `Confirm invoice ${item.invoice_number ?? ""} appears in Smart AP.`,
      }).eq("id", email_id);
      return json({ ok: true });
    }

    if (action === "complete") {
      if (!item.labeled_at) {
        const provider = await providerFor(db, mailbox);
        try {
          await provider.addLabel(item.provider_message_id, mailbox.processed_label);
        } catch (err) {
          console.warn("Labeling failed:", err);
        } finally {
          await provider.close?.();
        }
      }
      await db.from("emails").update({ status: "done", completed_at: now, completed_by: userId, labeled_at: item.labeled_at ?? now })
        .eq("id", email_id);
      return json({ ok: true });
    }

    if (action === "reopen") {
      await db.from("emails").update({ status: "action_needed", completed_at: null, completed_by: null }).eq("id", email_id);
      return json({ ok: true });
    }

    return json({ error: "action must be forward, complete or reopen" }, 400);
  } catch (err) {
    console.error(err);
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});
