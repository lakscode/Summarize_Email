-- Account statements: the classifier lists the invoice numbers on the statement (statement_invoices), and each is looked
-- up among the mailbox's processed invoices (_shared/statement.ts): forwarded (and when), received but not forwarded, or
-- not found. The result is kept on the email (statement_check, as of statement_checked_at) and shown in the Inbox, which
-- can check again (inbox/index.ts); email_processing_log keeps each processing's result. Written only by the edge
-- functions (service role).
alter table public.processed_emails
  add column if not exists statement_invoices text[] not null default '{}',
  add column if not exists statement_check jsonb,
  add column if not exists statement_checked_at timestamptz;

alter table public.email_processing_log add column if not exists statement_check jsonb;

-- Looking up invoices by mailbox and category.
create index if not exists processed_emails_mailbox_category_idx
  on public.processed_emails (mailbox_id, category) where invoice_number is not null;
