-- History of the forwarded-invoices reports (see *_forwarded_invoices_report.sql): one row per report run, scheduled or
-- "Send now", whether it was sent, had nothing to send, or failed, with the invoices it listed. Shown on the Reports
-- page. Written only by the edge functions (_shared/forward-report.ts); readable by whoever can open the mailbox (its
-- owner, or its organization; see *_organization_mailboxes.sql). Kept when the mailbox is disconnected.
create table if not exists public.reports (
  id uuid primary key default gen_random_uuid(),
  kind text not null default 'forwarded_invoices' check (kind in ('forwarded_invoices')),
  mailbox_id uuid references public.mailboxes (id) on delete set null,
  mailbox_address text not null,
  owner_user_id uuid not null references auth.users (id) on delete cascade,
  -- schedule: the daily report; manual: "Send now", by sent_by.
  trigger text not null check (trigger in ('schedule', 'manual')),
  sent_by uuid references auth.users (id) on delete set null,
  sent_by_name text,
  period_start timestamptz not null,
  period_end timestamptz not null,
  -- sent: emailed; empty: nothing forwarded, so not emailed; failed: see error.
  status text not null check (status in ('sent', 'empty', 'failed')),
  recipients text[] not null default '{}',
  subject text,
  item_count integer not null default 0,
  -- The invoices listed: [{subject, from, category, vendor, invoice_number, amount, due_date, forwarded_to, forwarded_at,
  -- web_link}].
  items jsonb not null default '[]',
  error text,
  created_at timestamptz not null default now()
);

create index if not exists reports_mailbox_idx on public.reports (mailbox_id, created_at desc);
create index if not exists reports_owner_idx on public.reports (owner_user_id, created_at desc);

alter table public.reports enable row level security;

drop policy if exists "Users can read reports of their mailboxes" on public.reports;
create policy "Users can read reports of their mailboxes"
  on public.reports for select to authenticated
  using ((select auth.uid()) = owner_user_id or (mailbox_id is not null and public.my_mailbox_role(mailbox_id) is not null));

revoke all on public.reports from anon, authenticated;
grant select on public.reports to authenticated;
