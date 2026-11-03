-- One version of the app again: v1 (web/v1, functions <name>-v1) was merged into the main one (web/, functions <name>).
-- Undoes *_organization_app_version.sql: the v1 scheduler's cron job and Vault secret, and the per-organization version.
select cron.unschedule(jobid) from cron.job where jobname = 'process-inbox-schedules-v1';

delete from vault.secrets where name = 'scheduler_v1_url';

alter table public.organizations drop constraint if exists organizations_app_version_check;
alter table public.organizations drop column if exists app_version;
