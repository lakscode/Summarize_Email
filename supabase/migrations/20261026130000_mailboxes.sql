-- Connected mailboxes. Created only by the `mailboxes` edge function (service role) after OAuth;
-- the Microsoft refresh token lives in Vault and never reaches the browser.
-- Written to be re-runnable: the hosted project may already have a mailboxes table from the
-- earlier AP-inbox schema, which is kept (with its data) and brought in line here.
create extension if not exists supabase_vault with schema vault;

create table if not exists public.mailboxes (
  id uuid primary key default gen_random_uuid(),
  owner_user_id uuid not null references auth.users (id) on delete cascade,
  provider text not null default 'microsoft' check (provider in ('microsoft')),
  email_address text not null,
  display_name text,
  credentials_secret_id uuid,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (owner_user_id, email_address)
);

alter table public.mailboxes add column if not exists updated_at timestamptz not null default now();

alter table public.mailboxes enable row level security;

drop policy if exists "own mailboxes" on public.mailboxes;
drop policy if exists "Users can read own mailboxes" on public.mailboxes;
drop policy if exists "Users can delete own mailboxes" on public.mailboxes;

create policy "Users can read own mailboxes"
  on public.mailboxes for select
  to authenticated
  using ((select auth.uid()) = owner_user_id);

create policy "Users can delete own mailboxes"
  on public.mailboxes for delete
  to authenticated
  using ((select auth.uid()) = owner_user_id);

-- Clients may read everything except the secret id, and may delete; inserts/updates go through the edge function.
revoke all on public.mailboxes from anon, authenticated;
grant select (id, owner_user_id, provider, email_address, display_name, active, created_at, updated_at)
  on public.mailboxes to authenticated;
grant delete on public.mailboxes to authenticated;

-- Vault helpers (service role only)
create or replace function public.mailbox_secret_put(p_secret_id uuid, p_secret text, p_name text)
returns uuid
language plpgsql security definer set search_path = '' as $$
begin
  if p_secret_id is null then
    return vault.create_secret(p_secret, p_name);
  end if;
  perform vault.update_secret(p_secret_id, p_secret);
  return p_secret_id;
end $$;

create or replace function public.mailbox_secret_get(p_secret_id uuid)
returns text
language sql security definer set search_path = '' as $$
  select decrypted_secret from vault.decrypted_secrets where id = p_secret_id
$$;

revoke all on function public.mailbox_secret_put(uuid, text, text) from public, anon, authenticated;
revoke all on function public.mailbox_secret_get(uuid) from public, anon, authenticated;
grant execute on function public.mailbox_secret_put(uuid, text, text) to service_role;
grant execute on function public.mailbox_secret_get(uuid) to service_role;

-- Remove the stored token from Vault when a mailbox is disconnected.
create or replace function public.mailboxes_delete_secret()
returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if old.credentials_secret_id is not null then
    delete from vault.secrets where id = old.credentials_secret_id;
  end if;
  return old;
end $$;

revoke all on function public.mailboxes_delete_secret() from public, anon, authenticated;

drop trigger if exists mailboxes_delete_secret on public.mailboxes;
create trigger mailboxes_delete_secret
  after delete on public.mailboxes
  for each row execute function public.mailboxes_delete_secret();
