-- One instance of each shared mailbox per organization: whoever in the organization added it first owns it, it is
-- shared with everyone in the organization, and nobody else in the organization can add it again (mailboxes/index.ts
-- refuses, pointing them to the existing one). Shared mailboxes: read through someone's own account (delegate_address)
-- or with organization API keys (auth_type 'app'). Personal mailboxes are only shared if their owner ticks it.
--   - Existing duplicates (same address, owners in the same organization) are merged into the first-added copy, which
--     is shared: emails it hasn't processed are moved over, the processing history points to it, and it takes the
--     duplicate's schedule if it had none. The duplicates are then deleted (their Vault tokens with them).
--   - Someone who joins (or is moved to) an organization: their shared mailboxes are shared with it, or merged into the
--     organization's copy when it already has one (replaces organization_members_share from
--     *_share_existing_shared_mailboxes.sql).
--   - A unique index keeps it that way: an organization can't have the same address shared twice.

-- Moves what `p_from` has into `p_into`, then deletes `p_from`.
create or replace function public.merge_mailbox_into(p_from uuid, p_into uuid)
returns void
language plpgsql security definer set search_path = '' as $$
declare
  v_owner uuid;
begin
  select owner_user_id into v_owner from public.mailboxes where id = p_into;
  -- Emails the kept copy hasn't processed (the same message processed by both stays with the kept copy's result).
  update public.processed_emails e set mailbox_id = p_into, owner_user_id = v_owner
    where e.mailbox_id = p_from
      and not exists (select 1 from public.processed_emails k where k.mailbox_id = p_into and k.message_id = e.message_id);
  update public.email_processing_log set mailbox_id = p_into, owner_user_id = v_owner where mailbox_id = p_from;
  -- Keep it running on a schedule if only the duplicate had one.
  update public.mailboxes k set
      schedule_enabled = d.schedule_enabled, schedule_frequency = d.schedule_frequency, schedule_days = d.schedule_days,
      schedule_time = d.schedule_time, schedule_timezone = d.schedule_timezone, schedule_anchor = d.schedule_anchor,
      schedule_interval = d.schedule_interval, schedule_interval_unit = d.schedule_interval_unit,
      schedule_next_run_at = d.schedule_next_run_at, schedule_last_run_at = d.schedule_last_run_at
    from public.mailboxes d
    where k.id = p_into and d.id = p_from and not k.schedule_enabled and d.schedule_enabled;
  -- Its roles and leftover emails go with it (cascade); its Vault token by trigger.
  delete from public.mailboxes where id = p_from;
end $$;

revoke all on function public.merge_mailbox_into(uuid, uuid) from public, anon, authenticated;

-- Existing data: per organization and address, the first-added copy stays and is shared.
do $$
declare
  r record;
begin
  for r in
    with scoped as (
      select m.id, o.org_id,
        first_value(m.id) over (partition by o.org_id, m.email_address order by m.created_at, m.id) as keep_id
      from public.mailboxes m
      join public.organization_members o on o.user_id = m.owner_user_id
      where m.delegate_address is not null or m.auth_type = 'app' or m.org_id = o.org_id
    )
    select id, keep_id, org_id from scoped
  loop
    if r.id <> r.keep_id then
      perform public.merge_mailbox_into(r.id, r.keep_id);
    end if;
  end loop;
end $$;

update public.mailboxes m set org_id = o.org_id, updated_at = now()
  from public.organization_members o
  where o.user_id = m.owner_user_id and m.org_id is distinct from o.org_id
    and (m.delegate_address is not null or m.auth_type = 'app');

create unique index if not exists mailboxes_org_address_key on public.mailboxes (org_id, email_address) where org_id is not null;

-- Joining an organization: share the newcomer's shared mailboxes with it, or merge them into its existing copy.
create or replace function public.organization_members_share()
returns trigger
language plpgsql security definer set search_path = '' as $$
declare
  r record;
  v_keep uuid;
begin
  if tg_op = 'UPDATE' and new.org_id = old.org_id then
    return null;
  end if;
  for r in
    select c.id, c.email_address from public.mailboxes c
    where c.owner_user_id = new.user_id and c.org_id is null
      and (c.delegate_address is not null or c.auth_type = 'app')
  loop
    select k.id into v_keep from public.mailboxes k
      where k.org_id = new.org_id and k.email_address = r.email_address and k.id <> r.id;
    if v_keep is null then
      update public.mailboxes set org_id = new.org_id, updated_at = now() where id = r.id;
    else
      perform public.merge_mailbox_into(r.id, v_keep);
    end if;
  end loop;
  return null;
end $$;

revoke all on function public.organization_members_share() from public, anon, authenticated;

drop trigger if exists organization_members_zz_share on public.organization_members;
create trigger organization_members_zz_share
  after insert or update of org_id on public.organization_members
  for each row execute function public.organization_members_share();
