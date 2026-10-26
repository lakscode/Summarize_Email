-- Remove a mailbox's stored credentials from Vault when the mailbox is disconnected.
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

create trigger mailboxes_delete_secret
  after delete on public.mailboxes
  for each row execute function public.mailboxes_delete_secret();
