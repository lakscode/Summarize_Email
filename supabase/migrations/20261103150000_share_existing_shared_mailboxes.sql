-- Shares the shared mailboxes added before they were shared by default (see mailboxes/index.ts, defaultOrgShare) with
-- their owner's organization, so everyone in it can see and use them (as members; see *_organization_mailboxes.sql).
-- Shared mailboxes: read through someone's own account (delegate_address) or with organization API keys (auth_type
-- 'app'); personal mailboxes stay private. When several people in an organization added the same one, one copy is
-- shared (the others stay with their owners), so it isn't listed or processed twice: the one with working access,
-- then active, then scheduled, then most recently processed. Mailboxes the organization already has shared are left
-- as they are. Owners can still stop sharing any of them on the Settings page.
with candidates as (
  select m.id, o.org_id,
    row_number() over (
      partition by o.org_id, m.email_address
      order by (m.credentials_secret_id is not null or m.connection_id is not null) desc,
        m.active desc,
        m.schedule_enabled desc,
        m.last_processed_at desc nulls last,
        m.updated_at desc
    ) as pick
  from public.mailboxes m
  join public.organization_members o on o.user_id = m.owner_user_id
  where m.org_id is null
    and (m.delegate_address is not null or m.auth_type = 'app')
    and not exists (
      select 1 from public.mailboxes s where s.org_id = o.org_id and s.email_address = m.email_address
    )
)
update public.mailboxes m
  set org_id = c.org_id, updated_at = now()
  from candidates c
  where c.id = m.id and c.pick = 1;

-- From now on, the same when someone joins (or is moved to) an organization: the shared mailboxes they already have are
-- shared with it, one copy per address, unless the organization already has that mailbox.
create or replace function public.organization_members_share()
returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if tg_op = 'UPDATE' and new.org_id = old.org_id then
    return null;
  end if;
  update public.mailboxes m set org_id = new.org_id, updated_at = now()
    where m.id in (
      select distinct on (c.email_address) c.id
      from public.mailboxes c
      where c.owner_user_id = new.user_id and c.org_id is null
        and (c.delegate_address is not null or c.auth_type = 'app')
        and not exists (select 1 from public.mailboxes s where s.org_id = new.org_id and s.email_address = c.email_address)
      order by c.email_address, (c.credentials_secret_id is not null or c.connection_id is not null) desc, c.active desc
    );
  return null;
end $$;

revoke all on function public.organization_members_share() from public, anon, authenticated;

-- Triggers on the same event fire in name order: this one must run after organization_members_unshare (which, on a
-- move, first takes the user's mailboxes out of the old organization).
drop trigger if exists organization_members_zz_share on public.organization_members;
create trigger organization_members_zz_share
  after insert or update of org_id on public.organization_members
  for each row execute function public.organization_members_share();
