-- Microsoft mailboxes connected with API keys (app-only / client credentials) instead of a user sign-in, typically
-- a shared mailbox. The Vault secret then holds JSON {tenant_id, client_id, client_secret} of the customer's own
-- Azure app registration (application permissions Mail.Read and Mail.Send, admin-consented) rather than a
-- refresh token.
alter table public.mailboxes add column if not exists auth_type text not null default 'oauth';
alter table public.mailboxes drop constraint if exists mailboxes_auth_type_check;
alter table public.mailboxes add constraint mailboxes_auth_type_check check (auth_type in ('oauth', 'app'));
grant select (auth_type) on public.mailboxes to authenticated;
