-- Gmail mailboxes: `provider` may now be 'google' as well as 'microsoft'. The refresh token is stored in Vault the
-- same way for both.
-- Projects that kept the earlier AP-inbox table have `provider` as the enum mail_provider (see *_mailboxes.sql), so
-- the value is added to the enum there; otherwise `provider` is text with a check constraint.
do $$
begin
  if exists (
    select 1 from pg_attribute
    where attrelid = 'public.mailboxes'::regclass and attname = 'provider' and atttypid = to_regtype('public.mail_provider')
  ) then
    alter type public.mail_provider add value if not exists 'google';
  else
    alter table public.mailboxes drop constraint if exists mailboxes_provider_check;
    alter table public.mailboxes add constraint mailboxes_provider_check check (provider in ('microsoft', 'google'));
  end if;
end $$;
