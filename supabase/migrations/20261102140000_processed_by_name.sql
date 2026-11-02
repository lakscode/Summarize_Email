-- The name of who last processed an email / mailbox (profile name, else email), kept as a snapshot like
-- processed_by_email, so the Inbox can show "date time · name" without reading other users' profiles.
alter table public.processed_emails add column if not exists processed_by_name text;
alter table public.email_processing_log add column if not exists processed_by_name text;
alter table public.mailboxes add column if not exists last_processed_by uuid references auth.users (id) on delete set null;
alter table public.mailboxes add column if not exists last_processed_by_name text;

grant select (last_processed_by, last_processed_by_name) on public.mailboxes to authenticated;

update public.processed_emails e
   set processed_by_name = coalesce(nullif(trim(p.full_name), ''), p.email)
  from public.profiles p
 where p.id = e.processed_by and e.processed_by_name is null;

update public.email_processing_log l
   set processed_by_name = coalesce(nullif(trim(p.full_name), ''), p.email)
  from public.profiles p
 where p.id = l.processed_by and l.processed_by_name is null;
