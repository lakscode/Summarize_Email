-- Someone in the Inbox can assign an email (a task) to a person with access to its mailbox (PATCH inbox {id, assign_to}).
-- Names are snapshots, like processed_by_name. assignment_seen_at is null until the assignee opens their "Assigned to
-- me" list: the sidebar counts those, and the first page after signing in lists them. Each assignment is also an
-- email_processing_log row (trigger 'user'). Written only by the `inbox` edge function (service role).
alter table public.processed_emails
  add column if not exists assigned_to uuid references auth.users (id) on delete set null,
  add column if not exists assigned_to_name text,
  add column if not exists assigned_by uuid references auth.users (id) on delete set null,
  add column if not exists assigned_by_name text,
  add column if not exists assigned_at timestamptz,
  add column if not exists assignment_seen_at timestamptz;

create index if not exists processed_emails_assigned_idx
  on public.processed_emails (assigned_to, assignment_seen_at) where assigned_to is not null;
