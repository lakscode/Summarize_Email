-- Where the run summary goes when send_summary is on. Null: an own mailbox gets it itself; a shared mailbox
-- (read through another account or an organization connection) sends it to the owner's account email.
alter table public.mailboxes add column if not exists summary_address text;

grant select (summary_address) on public.mailboxes to authenticated;
grant update (summary_address) on public.mailboxes to authenticated;
