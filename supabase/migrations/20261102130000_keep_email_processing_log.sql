-- Keep the processing history when its email or mailbox is removed (e.g. the mailbox is disconnected): the links
-- are cleared instead of the rows being deleted, and the subject and mailbox address are kept as snapshots.
-- Re-processing an email adds rows and never changes earlier ones. Rows go only with the owner's account.
alter table public.email_processing_log add column if not exists subject text;
alter table public.email_processing_log add column if not exists mailbox_address text;

alter table public.email_processing_log alter column processed_email_id drop not null;
alter table public.email_processing_log alter column mailbox_id drop not null;

alter table public.email_processing_log drop constraint if exists email_processing_log_processed_email_id_fkey;
alter table public.email_processing_log add constraint email_processing_log_processed_email_id_fkey
  foreign key (processed_email_id) references public.processed_emails (id) on delete set null;

alter table public.email_processing_log drop constraint if exists email_processing_log_mailbox_id_fkey;
alter table public.email_processing_log add constraint email_processing_log_mailbox_id_fkey
  foreign key (mailbox_id) references public.mailboxes (id) on delete set null;

update public.email_processing_log l
   set subject = e.subject, mailbox_address = m.email_address
  from public.processed_emails e
  join public.mailboxes m on m.id = e.mailbox_id
 where e.id = l.processed_email_id and l.subject is null;
