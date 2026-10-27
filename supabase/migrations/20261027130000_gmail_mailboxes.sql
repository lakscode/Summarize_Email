-- Gmail mailboxes: `provider` may now be 'google' as well as 'microsoft'. The refresh token is stored in Vault the
-- same way for both.
alter table public.mailboxes drop constraint if exists mailboxes_provider_check;
alter table public.mailboxes add constraint mailboxes_provider_check check (provider in ('microsoft', 'google'));
