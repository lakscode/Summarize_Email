-- The address a report was sent from: the mailbox itself, or (a shared mailbox without Send As in Exchange) the account
-- it is read through; see _shared/microsoft.ts sendFromMailbox. Shown on the Reports page.
alter table public.reports add column if not exists sent_from text;
