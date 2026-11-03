-- Daily report of the invoices (and utility bills) forwarded from a mailbox: at forward_report_time in
-- forward_report_timezone, the scheduler emails the list forwarded since the previous report (by the app or manually
-- in Outlook) to forward_report_addresses (comma-separated). Set on the Settings page by the mailbox's owner or admins
-- (mailboxes/index.ts); sent by the scheduler (_shared/forward-report.ts). Next / last run times are written only by
-- the edge functions.
alter table public.mailboxes
  add column if not exists forward_report_enabled boolean not null default false,
  add column if not exists forward_report_time time not null default '08:00',
  add column if not exists forward_report_timezone text not null default 'UTC',
  add column if not exists forward_report_addresses text,
  add column if not exists forward_report_next_at timestamptz,
  add column if not exists forward_report_last_at timestamptz,
  add column if not exists forward_report_last_result jsonb;

create index if not exists mailboxes_forward_report_due_idx
  on public.mailboxes (forward_report_next_at) where forward_report_enabled;

grant select (forward_report_enabled, forward_report_time, forward_report_timezone, forward_report_addresses,
  forward_report_next_at, forward_report_last_at, forward_report_last_result) on public.mailboxes to authenticated;
