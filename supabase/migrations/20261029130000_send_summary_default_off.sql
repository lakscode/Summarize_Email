-- New mailboxes start without the summary email; it is turned on per mailbox in its settings.
-- Existing mailboxes keep their current setting.
alter table public.mailboxes alter column send_summary set default false;
