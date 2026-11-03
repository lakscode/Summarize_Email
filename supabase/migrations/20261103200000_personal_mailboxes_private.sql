-- Personal mailboxes (a user's own, connected by signing in) are only ever visible to the person who connected them:
-- only shared mailboxes (read through someone's account, delegate_address, or with organization API keys, auth_type
-- 'app') can be organization mailboxes. Unshares any personal mailbox shared before and enforces it from now on, so
-- mailbox_role (and with it every RLS policy and the processing functions) gives others no access to them.
update public.mailboxes set org_id = null, updated_at = now()
where org_id is not null and delegate_address is null and auth_type <> 'app';

delete from public.mailbox_members mm
using public.mailboxes m
where m.id = mm.mailbox_id and m.org_id is null;

alter table public.mailboxes drop constraint if exists mailboxes_only_shared_in_organization;
alter table public.mailboxes add constraint mailboxes_only_shared_in_organization
  check (org_id is null or delegate_address is not null or auth_type = 'app');
