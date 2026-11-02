-- Who processed each email and when. processed_emails keeps the latest run's processor; email_processing_log keeps
-- every event: each time a run classifies the email (trigger 'manual' or 'schedule') and each time someone marks it
-- done or reopens it (trigger 'user'). processed_by_email is a snapshot so the history still reads after a user is
-- deleted. Written only by the `inbox` / `scheduler` edge functions (service role).
alter table public.processed_emails add column if not exists processed_by uuid references auth.users (id) on delete set null;
alter table public.processed_emails add column if not exists processed_by_email text;
alter table public.processed_emails add column if not exists processed_at timestamptz;

create table if not exists public.email_processing_log (
  id bigint generated always as identity primary key,
  processed_email_id uuid not null references public.processed_emails (id) on delete cascade,
  mailbox_id uuid not null references public.mailboxes (id) on delete cascade,
  owner_user_id uuid not null references auth.users (id) on delete cascade,
  processed_by uuid references auth.users (id) on delete set null,
  processed_by_email text,
  trigger text not null check (trigger in ('manual', 'schedule', 'user')),
  run_id text,
  status text not null,
  category text,
  action text,
  error text,
  processed_at timestamptz not null default now()
);

create index if not exists email_processing_log_email_idx
  on public.email_processing_log (processed_email_id, processed_at desc);
create index if not exists email_processing_log_owner_idx
  on public.email_processing_log (owner_user_id, processed_at desc);

alter table public.email_processing_log enable row level security;

drop policy if exists "Users can read own email processing log" on public.email_processing_log;
create policy "Users can read own email processing log"
  on public.email_processing_log for select
  to authenticated
  using ((select auth.uid()) = owner_user_id);

revoke all on public.email_processing_log from anon, authenticated;
grant select on public.email_processing_log to authenticated;
