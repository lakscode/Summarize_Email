// Organizations (a company's users; see migrations/*_organizations.sql). A user belongs to at most one.
// Super admins (migrations/*_super_admins.sql) create organizations and manage any of them; an organization's admins
// manage their own. Users join automatically by invitation, or by signing in with Microsoft at the organization's
// email domain.
//   GET                                   -> { super_admin, organizations?, organization, role, manage, members, invites }
//                                            for the caller's own organization (organization null if none), joining
//                                            one first if they aren't in one: an invitation for their email, else
//                                            (signed in with Microsoft) the organization whose domain matches.
//                                            Super admins also get `organizations`: all of them with member counts.
//   GET ?org_id=<id>                      -> super admins: the same for any organization
//   POST {action: "create", name, domain?}                 -> super admins: { organization }
//   POST {action: "list_users", org_id}                    -> super admins: { users: [{email, full_name, organization}] }
//                                                             everyone with an account who isn't in `org_id`, to pick
//                                                             from when adding people (organization: their current one)
//   Super admins, or admins of `org_id` (default: the caller's own organization):
//   POST {action: "update", org_id?, name?, domain?}       -> { organization } (domain: super admins only, as it decides
//                                                             who joins; "" turns automatic joining off)
//   POST {action: "add_members", org_id?, emails, role?}   -> { results: [{email, ok, status?, error?}] } each added
//                                                             (has an account), invited (joins on first visit) or,
//                                                             by a super admin, moved from another organization
//   POST {action: "set_role", org_id?, user_id, role}      -> { member }
//   POST {action: "remove_member", org_id?, user_id}       -> { removed } (any member may remove themselves: leave)
//   POST {action: "cancel_invite", org_id?, email}         -> { cancelled }
//   DELETE ?org_id=<id>                                    -> delete the organization (members are left in none)
// Writes use the service role after checking the caller's rights. Failures are written to app_logs.
import { corsHeaders, json } from "../_shared/cors.ts";
import { fail, logApi } from "../_shared/log.ts";
import { adminClient, callerUser } from "../_shared/supabase.ts";
import type { User } from "jsr:@supabase/supabase-js@2";

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const DOMAIN = /^[a-z0-9-]+(\.[a-z0-9-]+)+$/;
const ROLES = ["admin", "member"];
const MAX_BULK = 100;
// Anyone can have an address at these, so they can't decide who joins an organization.
const PUBLIC_DOMAINS = new Set([
  "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com", "yahoo.com", "icloud.com",
  "me.com", "aol.com", "proton.me", "protonmail.com", "gmx.com", "zoho.com",
]);

type Membership = { org_id: string; role: string };
// The caller's rights: super admin, and their own membership (if any).
type Caller = { user: User; superAdmin: boolean; own: Membership | null };

const emailDomain = (email: string) => email.split("@")[1] ?? "";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  const caller = await callerUser(req);
  if (!caller) return json({ error: "Not signed in" }, 401);
  const { user } = caller;
  const url = new URL(req.url);

  try {
    const c: Caller = { user, superAdmin: await isSuperAdmin(user.id), own: await membership(user.id) };

    if (req.method === "GET") {
      if (!c.own) c.own = await autoJoin(user);
      const orgId = url.searchParams.get("org_id");
      if (orgId && !c.superAdmin && orgId !== c.own?.org_id) return json({ error: "Organization not found" }, 404);
      return json({
        super_admin: c.superAdmin,
        ...(c.superAdmin ? { organizations: await allOrganizations() } : {}),
        ...await describe(c, orgId ?? c.own?.org_id ?? null),
      });
    }

    if (req.method === "POST") {
      const body = await req.json().catch(() => ({}));
      if (body.action === "create") {
        if (!c.superAdmin) return json({ error: "Only super admins can create organizations" }, 403);
        return await create(user.id, body);
      }
      if (body.action === "list_users") {
        if (!c.superAdmin) return json({ error: "Only super admins can list users" }, 403);
        return json({ users: await usersOutside(String(body.org_id ?? "")) });
      }
      if (body.action === "remove_member" && body.user_id === user.id && c.own) return await removeMember(user.id, c.own.org_id, user.id);
      const orgId = await managedOrg(c, body.org_id);
      if (orgId instanceof Response) return orgId;
      if (body.action === "update") {
        if (body.domain !== undefined && !c.superAdmin) return json({ error: "Only super admins can change the domain" }, 403);
        return await update(user.id, orgId, body);
      }
      if (body.action === "add_members") return await addMembers(c, orgId, body);
      if (body.action === "set_role") return await setRole(user.id, orgId, body);
      if (body.action === "remove_member") return await removeMember(user.id, orgId, String(body.user_id ?? ""));
      if (body.action === "cancel_invite") return await cancelInvite(user.id, orgId, body);
      return json({ error: "Unknown action" }, 400);
    }

    if (req.method === "DELETE") {
      const orgId = await managedOrg(c, url.searchParams.get("org_id"));
      if (orgId instanceof Response) return orgId;
      const { error } = await adminClient().from("organizations").delete().eq("id", orgId);
      if (error) return await fail("organizations/delete", error.message, 500, { userId: user.id, context: { org_id: orgId } });
      return json({ deleted: orgId });
    }

    return json({ error: "Method not allowed" }, 405);
  } catch (err) {
    await logApi("organizations", err, { userId: user.id, context: { method: req.method } });
    return json({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

async function isSuperAdmin(userId: string): Promise<boolean> {
  const { data, error } = await adminClient().from("super_admins").select("user_id").eq("user_id", userId).maybeSingle();
  if (error) throw new Error(error.message);
  return !!data;
}

async function membership(userId: string): Promise<Membership | null> {
  const { data, error } = await adminClient().from("organization_members").select("org_id, role").eq("user_id", userId).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

// The organization the caller may manage: `orgId` (default: their own) if they are a super admin or its admin.
async function managedOrg(c: Caller, orgId: unknown): Promise<string | Response> {
  const id = orgId ? String(orgId) : c.own?.org_id;
  if (!id) return json({ error: "You're not in an organization" }, 400);
  if (c.superAdmin) {
    const { data, error } = await adminClient().from("organizations").select("id").eq("id", id).maybeSingle();
    if (error) throw new Error(error.message);
    return data ? id : json({ error: "Organization not found" }, 404);
  }
  if (c.own?.org_id !== id || c.own.role !== "admin") return json({ error: "Only organization admins can do this" }, 403);
  return id;
}

// Puts a user who is in no organization into one: an invitation for their (confirmed) email first, else the
// organization whose domain is their email domain when they signed in with Microsoft (Azure verifies the address).
async function autoJoin(user: User): Promise<Membership | null> {
  const email = user.email?.toLowerCase();
  if (!email || !user.email_confirmed_at) return null;
  const db = adminClient();

  const { data: invite, error } = await db.from("organization_invites").select("org_id, role")
    .eq("email", email).order("created_at").limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  let join: Membership | null = invite;

  const providers: string[] = user.app_metadata?.providers ?? [user.app_metadata?.provider];
  if (!join && providers.includes("azure")) {
    const { data: org, error: orgError } = await db.from("organizations").select("id").eq("domain", emailDomain(email)).maybeSingle();
    if (orgError) throw new Error(orgError.message);
    // Someone removed from (or who left) the organization isn't joined again by domain.
    const { data: removed, error: removedError } = org
      ? await db.from("organization_removed").select("user_id").eq("org_id", org.id).eq("user_id", user.id).maybeSingle()
      : { data: null, error: null };
    if (removedError) throw new Error(removedError.message);
    if (org && !removed) join = { org_id: org.id, role: "member" };
  }
  if (!join) return null;

  // Another request may have joined meanwhile; the existing membership then stands.
  const { error: joinError } = await db.from("organization_members")
    .upsert({ user_id: user.id, org_id: join.org_id, role: join.role }, { onConflict: "user_id", ignoreDuplicates: true });
  if (joinError) throw new Error(joinError.message);
  if (invite) await db.from("organization_invites").delete().eq("org_id", invite.org_id).eq("email", email);
  await logApi("organizations/join", `Joined organization ${join.org_id} ${invite ? "by invitation" : "by email domain"}`, {
    userId: user.id, level: "info", context: { org_id: join.org_id, role: join.role },
  });
  return await membership(user.id);
}

async function allOrganizations() {
  const db = adminClient();
  const [orgs, members] = await Promise.all([
    db.from("organizations").select("id, name, domain, created_at").order("name"),
    db.from("organization_members").select("org_id"),
  ]);
  for (const r of [orgs, members]) if (r.error) throw new Error(r.error.message);
  const counts = new Map<string, number>();
  for (const m of members.data!) counts.set(m.org_id, (counts.get(m.org_id) ?? 0) + 1);
  return orgs.data!.map((o) => ({ ...o, member_count: counts.get(o.id) ?? 0 }));
}

const MAX_USERS = 2000;

// Users with an email who aren't in `orgId`, by email, with the name of the organization they are in (if any).
async function usersOutside(orgId: string) {
  const db = adminClient();
  const [profiles, members, orgs] = await Promise.all([
    db.from("profiles").select("id, email, full_name").not("email", "is", null).order("email").limit(MAX_USERS),
    db.from("organization_members").select("user_id, org_id"),
    db.from("organizations").select("id, name"),
  ]);
  for (const r of [profiles, members, orgs]) if (r.error) throw new Error(r.error.message);
  const orgOf = new Map(members.data!.map((m) => [m.user_id, m.org_id]));
  const orgName = new Map(orgs.data!.map((o) => [o.id, o.name]));
  return profiles.data!
    .filter((p) => orgOf.get(p.id) !== orgId)
    .map((p) => ({ email: p.email, full_name: p.full_name, organization: orgName.get(orgOf.get(p.id) ?? "") ?? null }));
}

// One organization as the caller sees it: `role` is theirs in it (null if not a member), `manage` whether they may
// change it. Invitations are listed only to those who manage it.
async function describe(c: Caller, orgId: string | null) {
  if (!orgId) return { organization: null };
  const db = adminClient();
  const [org, members, invites] = await Promise.all([
    db.from("organizations").select("id, name, domain, created_at").eq("id", orgId).maybeSingle(),
    db.from("organization_members").select("user_id, role, created_at").eq("org_id", orgId).order("created_at"),
    db.from("organization_invites").select("email, role, created_at").eq("org_id", orgId).order("created_at"),
  ]);
  for (const r of [org, members, invites]) if (r.error) throw new Error(r.error.message);
  if (!org.data) return { organization: null };
  const ids = members.data!.map((m) => m.user_id);
  const { data: profiles, error } = await db.from("profiles").select("id, email, full_name").in("id", ids);
  if (error) throw new Error(error.message);
  const byId = new Map((profiles ?? []).map((p) => [p.id, p]));
  const role = c.own?.org_id === orgId ? c.own.role : null;
  const manage = c.superAdmin || role === "admin";
  return {
    organization: org.data,
    role,
    manage,
    members: members.data!.map((m) => ({
      ...m, email: byId.get(m.user_id)?.email ?? null, full_name: byId.get(m.user_id)?.full_name ?? null, you: m.user_id === c.user.id,
    })),
    invites: manage ? invites.data : [],
  };
}

// A domain for automatic joining; not a public email service, where anyone could get an address.
function checkDomain(value: unknown): { domain: string | null } | { error: string } {
  const domain = String(value ?? "").trim().toLowerCase().replace(/^@/, "");
  if (!domain) return { domain: null };
  if (!DOMAIN.test(domain)) return { error: "Domain must look like yourcompany.com" };
  if (PUBLIC_DOMAINS.has(domain)) return { error: `${domain} is a public email service; anyone could join. Add people by email instead.` };
  return { domain };
}

const domainTaken = (message: string) => /organizations_domain_key|duplicate key/.test(message);

async function create(userId: string, body: Record<string, unknown>): Promise<Response> {
  const name = String(body.name ?? "").trim();
  if (!name) return json({ error: "Name is required" }, 400);
  const checked = checkDomain(body.domain);
  if ("error" in checked) return json({ error: checked.error }, 400);

  const { data: org, error } = await adminClient().from("organizations").insert({ name, domain: checked.domain, created_by: userId })
    .select("id, name, domain, created_at").single();
  if (error) {
    if (domainTaken(error.message)) return json({ error: `Another organization already uses ${checked.domain}` }, 400);
    return await fail("organizations/create", error.message, 500, { userId });
  }
  return json({ organization: org });
}

async function update(userId: string, orgId: string, body: Record<string, unknown>): Promise<Response> {
  const fields: Record<string, unknown> = { updated_at: new Date().toISOString() };
  if (body.name !== undefined) {
    const name = String(body.name ?? "").trim();
    if (!name) return json({ error: "Name is required" }, 400);
    fields.name = name;
  }
  if (body.domain !== undefined) {
    const checked = checkDomain(body.domain);
    if ("error" in checked) return json({ error: checked.error }, 400);
    fields.domain = checked.domain;
  }
  const { data, error } = await adminClient().from("organizations").update(fields).eq("id", orgId)
    .select("id, name, domain, created_at").single();
  if (error) {
    if (domainTaken(error.message)) return json({ error: `Another organization already uses ${fields.domain}` }, 400);
    return await fail("organizations/update", error.message, 500, { userId, context: { org_id: orgId } });
  }
  return json({ organization: data });
}

// Adds each address: people with an account join now (a super admin moves them from another organization), others
// are invited and join on their first visit.
async function addMembers(c: Caller, orgId: string, body: Record<string, unknown>): Promise<Response> {
  const raw = Array.isArray(body.emails) ? body.emails.join("\n") : String(body.emails ?? "");
  const emails = [...new Set(raw.split(/[\s,;]+/).map((a) => a.trim().toLowerCase()).filter(Boolean))];
  if (!emails.length) return json({ error: "Enter at least one email address" }, 400);
  if (emails.length > MAX_BULK) return json({ error: `Add at most ${MAX_BULK} people at a time` }, 400);
  const role = String(body.role ?? "member");
  if (!ROLES.includes(role)) return json({ error: "role must be admin or member" }, 400);

  const db = adminClient();
  const results: { email: string; ok: boolean; status?: "added" | "moved" | "invited"; error?: string }[] = [];
  for (const email of emails) {
    if (!EMAIL.test(email)) {
      results.push({ email, ok: false, error: "Not a valid email address" });
      continue;
    }
    // profiles.email keeps the case it was signed up with; escape LIKE wildcards so only this address matches.
    const { data: profile, error } = await db.from("profiles").select("id")
      .ilike("email", email.replace(/[\\%_]/g, "\\$&")).limit(1).maybeSingle();
    if (error) {
      results.push({ email, ok: false, error: error.message });
      continue;
    }
    if (!profile) {
      const { error: inviteError } = await db.from("organization_invites")
        .upsert({ org_id: orgId, email, role, invited_by: c.user.id }, { onConflict: "org_id,email" });
      results.push(inviteError ? { email, ok: false, error: inviteError.message } : { email, ok: true, status: "invited" });
      continue;
    }

    const existing = await membership(profile.id);
    if (existing?.org_id === orgId) {
      results.push({ email, ok: false, error: "Already a member" });
      continue;
    }
    if (existing && !c.superAdmin) {
      results.push({ email, ok: false, error: "Already in another organization" });
      continue;
    }
    if (existing?.role === "admin" && await adminCount(existing.org_id) <= 1) {
      results.push({ email, ok: false, error: "The only admin of another organization; make someone else admin there first" });
      continue;
    }
    const { error: addError } = await db.from("organization_members")
      .upsert({ user_id: profile.id, org_id: orgId, role, created_at: new Date().toISOString() }, { onConflict: "user_id" });
    if (addError) {
      results.push({ email, ok: false, error: addError.message });
      continue;
    }
    await db.from("organization_removed").delete().eq("org_id", orgId).eq("user_id", profile.id);
    if (existing) await db.from("organization_removed").upsert({ org_id: existing.org_id, user_id: profile.id }, { onConflict: "org_id,user_id" });
    results.push({ email, ok: true, status: existing ? "moved" : "added" });
  }
  return json({ results });
}

async function adminCount(orgId: string): Promise<number> {
  const { count, error } = await adminClient().from("organization_members").select("user_id", { count: "exact", head: true })
    .eq("org_id", orgId).eq("role", "admin");
  if (error) throw new Error(error.message);
  return count ?? 0;
}

const LAST_ADMIN = "An organization with members needs an admin. Make another member an admin first, or delete the organization.";

// True when taking `target` out of the admins would leave members with no admin (an organization emptied entirely
// is fine: super admins still manage it).
async function lastAdmin(orgId: string, target: Membership, newRole: string | null): Promise<boolean> {
  if (target.role !== "admin" || newRole === "admin") return false;
  if (await adminCount(orgId) > 1) return false;
  const { count, error } = await adminClient().from("organization_members").select("user_id", { count: "exact", head: true }).eq("org_id", orgId);
  if (error) throw new Error(error.message);
  return (count ?? 0) > 1;
}

async function setRole(userId: string, orgId: string, body: Record<string, unknown>): Promise<Response> {
  const role = String(body.role ?? "");
  if (!ROLES.includes(role)) return json({ error: "role must be admin or member" }, 400);
  const target = await membership(String(body.user_id ?? ""));
  if (!target || target.org_id !== orgId) return json({ error: "Member not found" }, 404);
  if (await lastAdmin(orgId, target, role)) return json({ error: LAST_ADMIN }, 400);
  const { data, error } = await adminClient().from("organization_members").update({ role })
    .eq("user_id", body.user_id).select("user_id, role").single();
  if (error) return await fail("organizations/set_role", error.message, 500, { userId });
  return json({ member: data });
}

async function removeMember(userId: string, orgId: string, targetId: string): Promise<Response> {
  const target = await membership(targetId);
  if (!target || target.org_id !== orgId) return json({ error: "Member not found" }, 404);
  if (await lastAdmin(orgId, target, null)) return json({ error: LAST_ADMIN }, 400);
  const db = adminClient();
  const { error } = await db.from("organization_members").delete().eq("user_id", targetId);
  if (error) return await fail("organizations/remove_member", error.message, 500, { userId });
  await db.from("organization_removed").upsert({ org_id: orgId, user_id: targetId }, { onConflict: "org_id,user_id" });
  return json({ removed: targetId });
}

async function cancelInvite(userId: string, orgId: string, body: Record<string, unknown>): Promise<Response> {
  const email = String(body.email ?? "").trim().toLowerCase();
  const { error, count } = await adminClient().from("organization_invites").delete({ count: "exact" })
    .eq("org_id", orgId).eq("email", email);
  if (error) return await fail("organizations/cancel_invite", error.message, 500, { userId });
  if (!count) return json({ error: "Invitation not found" }, 404);
  return json({ cancelled: email });
}
