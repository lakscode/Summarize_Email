-- Per-mailbox setting: email the run summary to the mailbox after each processing run. On by default, as before.
alter table public.mailboxes add column if not exists send_summary boolean not null default true;

grant select (send_summary) on public.mailboxes to authenticated;
grant update (send_summary) on public.mailboxes to authenticated;
