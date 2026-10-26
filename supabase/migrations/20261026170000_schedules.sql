-- Scheduled inbox processing: per-mailbox schedule, and a pg_cron job that calls the `scheduler` Edge Function
-- every 15 minutes. The function URL and a shared secret live in Vault; the secret is generated here, so no
-- manual secret setup is needed.
create extension if not exists pg_cron;
create extension if not exists pg_net with schema extensions;

alter table public.mailboxes add column if not exists schedule_enabled boolean not null default false;
alter table public.mailboxes add column if not exists schedule_frequency text not null default 'daily';
alter table public.mailboxes add column if not exists schedule_days smallint[] not null default '{1,2,3,4,5}';
alter table public.mailboxes add column if not exists schedule_time time not null default '08:00';
alter table public.mailboxes add column if not exists schedule_timezone text not null default 'UTC';
alter table public.mailboxes add column if not exists schedule_anchor date;
alter table public.mailboxes add column if not exists schedule_next_run_at timestamptz;
alter table public.mailboxes add column if not exists schedule_last_run_at timestamptz;
alter table public.mailboxes add column if not exists schedule_last_result jsonb;

alter table public.mailboxes drop constraint if exists mailboxes_schedule_frequency_check;
alter table public.mailboxes add constraint mailboxes_schedule_frequency_check
  check (schedule_frequency in ('daily', 'alternate', 'weekdays'));

create index if not exists mailboxes_schedule_due_idx
  on public.mailboxes (schedule_next_run_at) where schedule_enabled;

-- Readable by the owner; written only by the mailboxes / scheduler functions (service role).
grant select (schedule_enabled, schedule_frequency, schedule_days, schedule_time, schedule_timezone,
  schedule_next_run_at, schedule_last_run_at, schedule_last_result) on public.mailboxes to authenticated;

do $$
begin
  if not exists (select 1 from vault.secrets where name = 'cron_secret') then
    perform vault.create_secret(replace(gen_random_uuid()::text || gen_random_uuid()::text, '-', ''),
      'cron_secret', 'Shared secret between pg_cron and the scheduler Edge Function');
  end if;
  if not exists (select 1 from vault.secrets where name = 'scheduler_url') then
    -- Update this Vault secret if the project moves (Dashboard > Project Settings > Vault).
    perform vault.create_secret('https://bnzllbgsbrhiesaqnkcz.supabase.co/functions/v1/scheduler',
      'scheduler_url', 'URL of the scheduler Edge Function called by pg_cron');
  end if;
end $$;

-- Lets the scheduler function check the header it receives against the Vault secret.
create or replace function public.cron_secret()
returns text
language sql security definer set search_path = '' as $$
  select decrypted_secret from vault.decrypted_secrets where name = 'cron_secret'
$$;
revoke all on function public.cron_secret() from public, anon, authenticated;
grant execute on function public.cron_secret() to service_role;

select cron.unschedule(jobid) from cron.job where jobname = 'process-inbox-schedules';

select cron.schedule(
  'process-inbox-schedules',
  '*/15 * * * *',
  $job$
  select net.http_post(
    url := (select decrypted_secret from vault.decrypted_secrets where name = 'scheduler_url'),
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'cron_secret')
    ),
    body := '{}'::jsonb,
    timeout_milliseconds := 150000
  );
  $job$
);
