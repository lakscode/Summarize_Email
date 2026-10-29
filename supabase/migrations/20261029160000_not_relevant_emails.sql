-- Mail that isn't AP / AR at all (newsletters, notifications, internal chatter): category not_relevant, status ignored.
-- Ignored emails are kept so they aren't looked at again, but are left out of the Inbox lists and the summary email.
alter table public.processed_emails drop constraint if exists processed_emails_category_check;
alter table public.processed_emails add constraint processed_emails_category_check check (category in (
  'invoice', 'utility_bill', 'vendor_info', 'statement',
  'remittance', 'purchase_order', 'customer_inquiry',
  'other', 'not_relevant'
));

alter table public.processed_emails drop constraint if exists processed_emails_status_check;
alter table public.processed_emails add constraint processed_emails_status_check
  check (status in ('processing', 'action_needed', 'done', 'no_action', 'ignored', 'error'));

-- Only invoices, utility bills, vendor info and statements (AP) / remittances, orders and inquiries (AR) are processed
-- now; "other" is no longer a category. Emails already filed as other move to not relevant.
update public.processed_emails
set category = 'not_relevant', status = 'ignored', updated_at = now()
where category = 'other';
