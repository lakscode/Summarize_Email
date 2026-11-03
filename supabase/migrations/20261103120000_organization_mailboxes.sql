-- Organization mailboxes (v1): the user who connected a mailbox may share it with their organization. Everyone in the
-- organization can then use it with a role per mailbox:
--   admin  - change its settings and schedule, and who is an admin of it
--   member - see it, its emails and history, process its inbox and mark emails done (the default)
-- The owner (who connected it; its token is theirs) and organization admins are always admin. mailbox_members keeps
-- only roles set by hand, so people who join the organization later get access as members. A mailbox whose owner is no
-- longer in the organization is not shared with it any more (see mailbox_role).
alter table public.mailboxes add column if not exists org_id uuid references public.organizations (id) on delete set null;
create index if not exists mailboxes_org_idx on public.mailboxes (org_id) where org_id is not null;
grant select (org_id) on public.mailboxes to authenticated;

create table if not exists public.mailbox_members (
  mailbox_id uuid not null references public.mailboxes (id) on delete cascade,
  user_id uuid not null references auth.users (id) on delete cascade,
  role text not null default 'member' check (role in ('admin', 'member')),
  updated_by uuid references auth.users (id) on delete set null,
  updated_at timestamptz not null default now(),
  primary key (mailbox_id, user_id)
);

-- 'owner', 'admin', 'member', or null (no access) for `p_user_id` on a mailbox. Service role only: the edge functions
-- check rights with it before writing with the service role.
create or replace function public.mailbox_role(p_mailbox_id uuid, p_user_id uuid)
returns text
language sql stable security definer set search_path = '' as $$
  select case
    when m.owner_user_id = p_user_id then 'owner'
    when m.org_id is null or me.org_id is distinct from m.org_id then null
    when not exists (select 1 from public.organization_members o where o.user_id = m.owner_user_id and o.org_id = m.org_id) then null
    when me.role = 'admin' then 'admin'
    else coalesce(mm.role, 'member')
  end
  from public.mailboxes m
  left join public.organization_members me on me.user_id = p_user_id
  left join public.mailbox_members mm on mm.mailbox_id = m.id and mm.user_id = p_user_id
  where m.id = p_mailbox_id
$$;

-- The caller's role, for RLS.
create or replace function public.my_mailbox_role(p_mailbox_id uuid)
returns text
language sql stable security definer set search_path = '' as $$
  select public.mailbox_role(p_mailbox_id, (select auth.uid()))
$$;

revoke all on function public.mailbox_role(uuid, uuid) from public, anon, authenticated;
grant execute on function public.mailbox_role(uuid, uuid) to service_role;
revoke all on function public.my_mailbox_role(uuid) from public, anon;
grant execute on function public.my_mailbox_role(uuid) to authenticated;

-- Read access for everyone in the organization, alongside the owner policies. Writes to mailboxes and mailbox_members
-- go through the edge functions (service role).
drop policy if exists "Organization can read shared mailboxes" on public.mailboxes;
create policy "Organization can read shared mailboxes"
  on public.mailboxes for select to authenticated
  using (org_id is not null and public.my_mailbox_role(id) is not null);

drop policy if exists "Organization can read shared mailbox emails" on public.processed_emails;
create policy "Organization can read shared mailbox emails"
  on public.processed_emails for select to authenticated
  using (public.my_mailbox_role(mailbox_id) is not null);

-- Members mark emails done / reopen them.
drop policy if exists "Organization can update shared mailbox emails" on public.processed_emails;
create policy "Organization can update shared mailbox emails"
  on public.processed_emails for update to authenticated
  using (public.my_mailbox_role(mailbox_id) is not null)
  with check (public.my_mailbox_role(mailbox_id) is not null);

drop policy if exists "Organization can read shared mailbox history" on public.email_processing_log;
create policy "Organization can read shared mailbox history"
  on public.email_processing_log for select to authenticated
  using (mailbox_id is not null and public.my_mailbox_role(mailbox_id) is not null);

alter table public.mailbox_members enable row level security;
drop policy if exists "Organization can read shared mailbox roles" on public.mailbox_members;
create policy "Organization can read shared mailbox roles"
  on public.mailbox_members for select to authenticated
  using (public.my_mailbox_role(mailbox_id) is not null);
revoke all on public.mailbox_members from anon, authenticated;
grant select on public.mailbox_members to authenticated;
