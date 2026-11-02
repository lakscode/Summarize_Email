-- Each email_processing_log row keeps a full snapshot of the email as that processing left it (sender, dates, what
-- was extracted, the summary and next step), so the record stays complete after the email is processed again
-- (processed_emails keeps only the latest) or its mailbox is disconnected (processed_emails rows are deleted).
alter table public.email_processing_log
  add column if not exists message_id text,
  add column if not exists received_at timestamptz,
  add column if not exists from_name text,
  add column if not exists from_address text,
  add column if not exists web_link text,
  add column if not exists attachment_names text[] not null default '{}',
  add column if not exists confidence text,
  add column if not exists vendor text,
  add column if not exists invoice_number text,
  add column if not exists amount text,
  add column if not exists due_date text,
  add column if not exists summary text,
  add column if not exists next_step text;

create index if not exists email_processing_log_mailbox_idx
  on public.email_processing_log (mailbox_address, processed_at desc);

-- Existing rows: the email's current details (the closest there is to what they were then).
update public.email_processing_log l
   set message_id = e.message_id,
       received_at = e.received_at,
       from_name = e.from_name,
       from_address = e.from_address,
       web_link = e.web_link,
       attachment_names = e.attachment_names,
       confidence = e.confidence,
       vendor = e.vendor,
       invoice_number = e.invoice_number,
       amount = e.amount,
       due_date = e.due_date,
       summary = e.summary,
       next_step = e.next_step
  from public.processed_emails e
 where e.id = l.processed_email_id and l.message_id is null;
