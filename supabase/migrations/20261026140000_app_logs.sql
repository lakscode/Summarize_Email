-- Error/warning log for both the UI (source 'ui', written by the browser through the `logs` function)
-- and the Edge Functions (source 'api', written with the service role). Users see only their own rows.
create table if not exists public.app_logs (
  id bigint generated always as identity primary key,
  user_id uuid references auth.users (id) on delete cascade,
  source text not null check (source in ('ui', 'api')),
  level text not null default 'error' check (level in ('error', 'warn', 'info')),
  location text,
  message text not null,
  context jsonb not null default '{}',
  created_at timestamptz not null default now()
);

create index if not exists app_logs_user_created_idx on public.app_logs (user_id, created_at desc);

alter table public.app_logs enable row level security;

drop policy if exists "Users can read own logs" on public.app_logs;
drop policy if exists "Users can write own ui logs" on public.app_logs;
drop policy if exists "Users can clear own logs" on public.app_logs;

create policy "Users can read own logs"
  on public.app_logs for select
  to authenticated
  using ((select auth.uid()) = user_id);

create policy "Users can write own ui logs"
  on public.app_logs for insert
  to authenticated
  with check ((select auth.uid()) = user_id and source = 'ui');

create policy "Users can clear own logs"
  on public.app_logs for delete
  to authenticated
  using ((select auth.uid()) = user_id);

revoke all on public.app_logs from anon, authenticated;
grant select, delete on public.app_logs to authenticated;
grant insert (user_id, source, level, location, message, context) on public.app_logs to authenticated;
