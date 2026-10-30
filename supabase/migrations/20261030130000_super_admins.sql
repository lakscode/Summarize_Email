-- App-wide administrators ("super admins"): the only users who create organizations, and they can manage any of
-- them (rename, set the auto-join domain, assign and remove users, delete). Organization admins still manage their
-- own organization's members. There is no UI to grant this; add one by hand (SQL editor):
--   insert into public.super_admins (user_id) select id from auth.users where email = 'you@yourcompany.com';
create table if not exists public.super_admins (
  user_id uuid primary key references auth.users (id) on delete cascade,
  created_at timestamptz not null default now()
);

-- No policies: read only by the edge functions (service role).
alter table public.super_admins enable row level security;
revoke all on public.super_admins from anon, authenticated;
