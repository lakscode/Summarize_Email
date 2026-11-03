-- App version per organization: 'current' (web/, functions <name>) or 'v1' (web/v1/, functions <name>-v1). Both run
-- side by side on the same database. Each UI sends members of an organization on the other version to that version's
-- pages, and each scheduler only runs mailboxes whose owner is on its version (users in no organization: current).
-- Set by super admins on the Organization page.
alter table public.organizations add column if not exists app_version text not null default 'current';

alter table public.organizations drop constraint if exists organizations_app_version_check;
alter table public.organizations add constraint organizations_app_version_check
  check (app_version in ('current', 'v1'));

-- pg_cron calls the v1 scheduler as well (same shared secret). Update this Vault secret if the project moves.
do $$
begin
  if not exists (select 1 from vault.secrets where name = 'scheduler_v1_url') then
    perform vault.create_secret('https://bnzllbgsbrhiesaqnkcz.supabase.co/functions/v1/scheduler-v1',
      'scheduler_v1_url', 'URL of the v1 scheduler Edge Function called by pg_cron');
  end if;
end $$;

select cron.unschedule(jobid) from cron.job where jobname = 'process-inbox-schedules-v1';

select cron.schedule(
  'process-inbox-schedules-v1',
  '*/5 * * * *',
  $job$
  select net.http_post(
    url := (select decrypted_secret from vault.decrypted_secrets where name = 'scheduler_v1_url'),
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'cron_secret')
    ),
    body := '{}'::jsonb,
    timeout_milliseconds := 150000
  );
  $job$
);
