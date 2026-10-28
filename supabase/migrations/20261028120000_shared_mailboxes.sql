-- Microsoft 365 shared mailboxes: the user signs in with their own account and the app reads the shared mailbox
-- through it (delegated Mail.Read.Shared / Mail.Send.Shared). email_address is the shared mailbox; delegate_address
-- is the account that signed in and whose refresh token is stored. Null for a user's own mailbox.
alter table public.mailboxes add column if not exists delegate_address text;
grant select (delegate_address) on public.mailboxes to authenticated;
