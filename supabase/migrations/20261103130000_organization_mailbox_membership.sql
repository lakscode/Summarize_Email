-- Organization mailboxes are only for the people of the organization they are shared with (see
-- *_organization_mailboxes.sql). mailbox_role() already denies everyone else; this keeps the data in line when someone
-- leaves an organization or is moved to another one (by either app version, or by hand):
--   - the mailboxes they shared with the old organization stop being shared (they keep them as their own), and
--   - the roles set on those mailboxes, and their own roles on the old organization's other mailboxes, are removed.
-- Otherwise a moved user's mailbox would still look shared on their Settings page, now naming the new organization.
create or replace function public.organization_members_unshare()
returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if tg_op = 'UPDATE' and new.org_id = old.org_id then
    return null;
  end if;
  delete from public.mailbox_members mm
    using public.mailboxes m
    where mm.mailbox_id = m.id and m.org_id = old.org_id
      and (m.owner_user_id = old.user_id or mm.user_id = old.user_id);
  update public.mailboxes set org_id = null, updated_at = now()
    where owner_user_id = old.user_id and org_id = old.org_id;
  return null;
end $$;

revoke all on function public.organization_members_unshare() from public, anon, authenticated;

drop trigger if exists organization_members_unshare on public.organization_members;
create trigger organization_members_unshare
  after delete or update of org_id on public.organization_members
  for each row execute function public.organization_members_unshare();

-- Clean up anything left from before this trigger: roles of people no longer in the mailbox's organization, then
-- mailboxes whose owner is no longer in the organization they were shared with.
delete from public.mailbox_members mm
  using public.mailboxes m
  where mm.mailbox_id = m.id
    and (m.org_id is null
      or not exists (select 1 from public.organization_members o where o.user_id = mm.user_id and o.org_id = m.org_id)
      or not exists (select 1 from public.organization_members o where o.user_id = m.owner_user_id and o.org_id = m.org_id));

update public.mailboxes m set org_id = null, updated_at = now()
  where m.org_id is not null
    and not exists (select 1 from public.organization_members o where o.user_id = m.owner_user_id and o.org_id = m.org_id);
