-- Organization connections and per-mailbox AP/AR settings, for running many (~50) shared AP/AR mailboxes.
--
-- ms_connections: one Microsoft 365 tenant's API keys (the customer's own Azure app registration with application
-- permissions Mail.Read and Mail.Send). Entered once; shared mailboxes are then added under it by address. The
-- keys live in Vault as JSON {tenant_id, client_id, client_secret}. Mailboxes added under a connection have
-- auth_type 'app', connection_id set, and no credentials of their own.
create table if not exists public.ms_connections (
  id uuid primary key default gen_random_uuid(),
  owner_user_id uuid not null references auth.users (id) on delete cascade,
  name text not null,
  tenant_id text not null,
  client_id text not null,
  credentials_secret_id uuid,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.ms_connections enable row level security;

drop policy if exists "Users can read own connections" on public.ms_connections;
drop policy if exists "Users can delete own connections" on public.ms_connections;
create policy "Users can read own connections"
  on public.ms_connections for select to authenticated using ((select auth.uid()) = owner_user_id);
create policy "Users can delete own connections"
  on public.ms_connections for delete to authenticated using ((select auth.uid()) = owner_user_id);

-- Clients may read everything except the secret id, and may delete; inserts/updates go through the edge function.
revoke all on public.ms_connections from anon, authenticated;
grant select (id, owner_user_id, name, tenant_id, client_id, created_at, updated_at) on public.ms_connections to authenticated;
grant delete on public.ms_connections to authenticated;

-- Remove the stored keys from Vault when a connection is deleted (its mailboxes go with it, see connection_id).
create or replace function public.ms_connections_delete_secret()
returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if old.credentials_secret_id is not null then
    delete from vault.secrets where id = old.credentials_secret_id;
  end if;
  return old;
end $$;

revoke all on function public.ms_connections_delete_secret() from public, anon, authenticated;

drop trigger if exists ms_connections_delete_secret on public.ms_connections;
create trigger ms_connections_delete_secret
  after delete on public.ms_connections
  for each row execute function public.ms_connections_delete_secret();

-- Drops a mailbox's own sign-in token when it moves under an organization connection.
create or replace function public.mailbox_secret_delete(p_secret_id uuid)
returns void
language sql security definer set search_path = '' as $$
  delete from vault.secrets where id = p_secret_id
$$;

revoke all on function public.mailbox_secret_delete(uuid) from public, anon, authenticated;
grant execute on function public.mailbox_secret_delete(uuid) to service_role;

alter table public.mailboxes add column if not exists connection_id uuid references public.ms_connections (id) on delete cascade;
create index if not exists mailboxes_connection_idx on public.mailboxes (connection_id);

-- Per-mailbox settings: AP or AR (which categories the classifier uses), and which categories are forwarded to the
-- mailbox's destination address (smart_ap_address).
alter table public.mailboxes add column if not exists mailbox_type text not null default 'ap';
alter table public.mailboxes drop constraint if exists mailboxes_mailbox_type_check;
alter table public.mailboxes add constraint mailboxes_mailbox_type_check check (mailbox_type in ('ap', 'ar'));
alter table public.mailboxes add column if not exists forward_categories text[] not null default '{invoice}';

grant select (connection_id, mailbox_type, forward_categories) on public.mailboxes to authenticated;
grant update (mailbox_type, forward_categories) on public.mailboxes to authenticated;

-- AR categories alongside the AP ones.
alter table public.processed_emails drop constraint if exists processed_emails_category_check;
alter table public.processed_emails add constraint processed_emails_category_check check (category in (
  'invoice', 'utility_bill', 'vendor_info', 'statement',
  'remittance', 'purchase_order', 'customer_inquiry',
  'other'
));
