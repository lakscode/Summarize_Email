-- AP inbox processing: per-mailbox Smart AP forwarding address, and one row per processed email.
alter table public.mailboxes add column if not exists smart_ap_address text;
alter table public.mailboxes add column if not exists last_processed_at timestamptz;

grant select (smart_ap_address, last_processed_at) on public.mailboxes to authenticated;
grant update (smart_ap_address, updated_at) on public.mailboxes to authenticated;

drop policy if exists "Users can update own mailboxes" on public.mailboxes;
create policy "Users can update own mailboxes"
  on public.mailboxes for update
  to authenticated
  using ((select auth.uid()) = owner_user_id)
  with check ((select auth.uid()) = owner_user_id);

-- Written by the `inbox` edge function (service role). Users read their rows and may change status.
create table if not exists public.processed_emails (
  id uuid primary key default gen_random_uuid(),
  mailbox_id uuid not null references public.mailboxes (id) on delete cascade,
  owner_user_id uuid not null references auth.users (id) on delete cascade,
  message_id text not null,
  received_at timestamptz,
  from_name text,
  from_address text,
  subject text,
  web_link text,
  attachment_names text[] not null default '{}',
  category text check (category in ('invoice', 'utility_bill', 'vendor_info', 'statement', 'other')),
  confidence text,
  vendor text,
  invoice_number text,
  amount text,
  due_date text,
  summary text,
  next_step text,
  action_taken text,
  -- processing: claimed by a run; action_needed: employee must act; done: handled; no_action; error: see error
  status text not null default 'processing'
    check (status in ('processing', 'action_needed', 'done', 'no_action', 'error')),
  error text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (mailbox_id, message_id)
);

create index if not exists processed_emails_owner_idx
  on public.processed_emails (owner_user_id, status, received_at desc);

alter table public.processed_emails enable row level security;

drop policy if exists "Users can read own processed emails" on public.processed_emails;
drop policy if exists "Users can update own processed emails" on public.processed_emails;

create policy "Users can read own processed emails"
  on public.processed_emails for select
  to authenticated
  using ((select auth.uid()) = owner_user_id);

create policy "Users can update own processed emails"
  on public.processed_emails for update
  to authenticated
  using ((select auth.uid()) = owner_user_id)
  with check ((select auth.uid()) = owner_user_id);

revoke all on public.processed_emails from anon, authenticated;
grant select on public.processed_emails to authenticated;
grant update (status, updated_at) on public.processed_emails to authenticated;
