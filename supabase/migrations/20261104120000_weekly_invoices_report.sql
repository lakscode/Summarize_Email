-- Weekly report of the invoices (and utility bills) forwarded from a mailbox, separate from the daily one
-- (*_forwarded_invoices_report.sql): on weekly_report_day (0 = Sunday ... 6 = Saturday) at weekly_report_time in
-- weekly_report_timezone, the scheduler emails the list forwarded in the 7 days before to weekly_report_addresses
-- (comma-separated). Set on the Settings page by the mailbox's owner or admins (mailboxes/index.ts); sent by the
-- scheduler (_shared/weekly-report.ts). Next / last run times are written only by the edge functions.
alter table public.mailboxes
  add column if not exists weekly_report_enabled boolean not null default false,
  add column if not exists weekly_report_day smallint not null default 1 check (weekly_report_day between 0 and 6),
  add column if not exists weekly_report_time time not null default '08:00',
  add column if not exists weekly_report_timezone text not null default 'UTC',
  add column if not exists weekly_report_addresses text,
  add column if not exists weekly_report_next_at timestamptz,
  add column if not exists weekly_report_last_at timestamptz,
  add column if not exists weekly_report_last_result jsonb;

create index if not exists mailboxes_weekly_report_due_idx
  on public.mailboxes (weekly_report_next_at) where weekly_report_enabled;

grant select (weekly_report_enabled, weekly_report_day, weekly_report_time, weekly_report_timezone, weekly_report_addresses,
  weekly_report_next_at, weekly_report_last_at, weekly_report_last_result) on public.mailboxes to authenticated;

-- Weekly reports go in the same report history (Reports page), told apart by kind.
alter table public.reports drop constraint if exists reports_kind_check;
alter table public.reports add constraint reports_kind_check check (kind in ('forwarded_invoices', 'weekly_invoices'));
