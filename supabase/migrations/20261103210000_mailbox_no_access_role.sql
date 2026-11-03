-- "No access" role on an organization mailbox: the mailbox's admins can hide it from someone in the organization.
-- mailbox_role() then returns null for them, so RLS (mailboxes, processed_emails, email_processing_log, reports) and the
-- functions treat the mailbox as not theirs. The owner and organization admins always keep access.
alter table public.mailbox_members drop constraint if exists mailbox_members_role_check;
alter table public.mailbox_members add constraint mailbox_members_role_check check (role in ('admin', 'member', 'none'));

create or replace function public.mailbox_role(p_mailbox_id uuid, p_user_id uuid)
returns text
language sql stable security definer set search_path = '' as $$
  select case
    when m.owner_user_id = p_user_id then 'owner'
    when m.org_id is null or me.org_id is distinct from m.org_id then null
    when not exists (select 1 from public.organization_members o where o.user_id = m.owner_user_id and o.org_id = m.org_id) then null
    when me.role = 'admin' then 'admin'
    when mm.role = 'none' then null
    else coalesce(mm.role, 'member')
  end
  from public.mailboxes m
  left join public.organization_members me on me.user_id = p_user_id
  left join public.mailbox_members mm on mm.mailbox_id = m.id and mm.user_id = p_user_id
  where m.id = p_mailbox_id
$$;
