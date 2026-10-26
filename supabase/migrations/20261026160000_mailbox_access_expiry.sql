-- Time-limited mailbox access: the user picks how many days the app may use the mailbox when connecting.
-- After access_expires_at the stored refresh token is deleted from Vault and the mailbox must be reconnected.
alter table public.mailboxes add column if not exists access_expires_at timestamptz;
grant select (access_expires_at) on public.mailboxes to authenticated;

-- Reuse a Vault secret with the same name when the mailbox row lost its pointer to it (e.g. a connect
-- that stored the token but failed before linking it), instead of failing on the unique name.
create or replace function public.mailbox_secret_put(p_secret_id uuid, p_secret text, p_name text)
returns uuid
language plpgsql security definer set search_path = '' as $$
declare
  v_id uuid := p_secret_id;
begin
  if v_id is null then
    select id into v_id from vault.secrets where name = p_name;
  end if;
  if v_id is null then
    return vault.create_secret(p_secret, p_name);
  end if;
  perform vault.update_secret(v_id, p_secret);
  return v_id;
end $$;

revoke all on function public.mailbox_secret_put(uuid, text, text) from public, anon, authenticated;
grant execute on function public.mailbox_secret_put(uuid, text, text) to service_role;

-- Deletes stored tokens for every mailbox whose access period has ended. Returns how many were expired.
create or replace function public.expire_mailbox_access()
returns integer
language plpgsql security definer set search_path = '' as $$
declare
  v_count integer;
begin
  delete from vault.secrets s
  using public.mailboxes m
  where m.credentials_secret_id = s.id
    and m.access_expires_at is not null
    and m.access_expires_at <= now();

  update public.mailboxes
  set credentials_secret_id = null, active = false, updated_at = now()
  where access_expires_at is not null
    and access_expires_at <= now()
    and (credentials_secret_id is not null or active);
  get diagnostics v_count = row_count;
  return v_count;
end $$;

revoke all on function public.expire_mailbox_access() from public, anon, authenticated;
grant execute on function public.expire_mailbox_access() to service_role;
