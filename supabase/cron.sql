-- Schedule the AP inbox run. Run this once in the SQL editor of your Supabase project
-- (it is not a migration because it contains your project URL and cron secret).
--
-- Replace <project-ref> and <CRON_SECRET> (the same value you set with `supabase secrets set`).

create extension if not exists pg_cron;
create extension if not exists pg_net;

select vault.create_secret('https://<project-ref>.supabase.co', 'ap_project_url');
select vault.create_secret('<CRON_SECRET>', 'ap_cron_secret');

-- Weekdays at 07:00 UTC. Adjust the cron expression to your timezone / cadence.
select cron.schedule(
  'ap-inbox-daily',
  '0 7 * * 1-5',
  $$
  select net.http_post(
    url     := (select decrypted_secret from vault.decrypted_secrets where name = 'ap_project_url')
               || '/functions/v1/process-inbox',
    headers := jsonb_build_object(
                 'Content-Type', 'application/json',
                 'x-cron-secret', (select decrypted_secret from vault.decrypted_secrets where name = 'ap_cron_secret')),
    body    := '{}'::jsonb,
    timeout_milliseconds := 10000
  );
  $$
);

-- To change or remove later:
--   select cron.unschedule('ap-inbox-daily');
--   select * from cron.job_run_details order by start_time desc limit 20;
