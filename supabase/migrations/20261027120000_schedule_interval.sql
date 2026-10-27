-- Interval schedules: "every N minutes / hours / days" counted from a start date (schedule_anchor) and time
-- (schedule_time). The cron tick moves from every 15 to every 5 minutes so short intervals run close to on time.
alter table public.mailboxes add column if not exists schedule_interval integer not null default 1;
alter table public.mailboxes add column if not exists schedule_interval_unit text not null default 'hours';

alter table public.mailboxes drop constraint if exists mailboxes_schedule_frequency_check;
alter table public.mailboxes add constraint mailboxes_schedule_frequency_check
  check (schedule_frequency in ('daily', 'alternate', 'weekdays', 'interval'));

alter table public.mailboxes drop constraint if exists mailboxes_schedule_interval_check;
alter table public.mailboxes add constraint mailboxes_schedule_interval_check
  check (schedule_interval between 1 and 1000 and schedule_interval_unit in ('minutes', 'hours', 'days'));

grant select (schedule_interval, schedule_interval_unit, schedule_anchor) on public.mailboxes to authenticated;

select cron.alter_job(jobid, schedule := '*/5 * * * *') from cron.job where jobname = 'process-inbox-schedules';
