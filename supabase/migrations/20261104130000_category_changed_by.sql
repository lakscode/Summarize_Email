-- Someone in the Inbox can correct an email's category (PATCH inbox {id, category}). Who did it and when are kept on the
-- email (a name snapshot, like processed_by_name) so a re-check run leaves the correction alone (see process.ts
-- recheckable); each change is also an email_processing_log row (trigger 'user'). Written only by the `inbox` edge
-- function (service role).
alter table public.processed_emails
  add column if not exists category_changed_by uuid references auth.users (id) on delete set null,
  add column if not exists category_changed_by_name text,
  add column if not exists category_changed_at timestamptz;
