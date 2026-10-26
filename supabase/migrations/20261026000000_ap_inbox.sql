-- AP Inbox: mailbox connections, processed emails, statement reconciliation, AP ledger.

create extension if not exists supabase_vault with schema vault;

-- ---------------------------------------------------------------------------
-- Enums
-- ---------------------------------------------------------------------------
create type public.mail_provider as enum ('microsoft', 'microsoft_app', 'gmail', 'imap');

create type public.email_category as enum (
  'invoice',        -- A. payable invoice -> Smart AP
  'utility_bill',   -- B. portal-only bill -> employee downloads from portal
  'vendor_info',    -- C. information only -> employee reviews
  'statement',      -- D. account statement -> reconcile, list missing invoices
  'other'           -- E. needs attention / not AP
);

create type public.item_status as enum (
  'done',           -- routine action completed, nothing for the employee
  'action_needed',  -- employee must do something
  'flagged',        -- possible fraud / needs human verification
  'info',           -- no action
  'skipped'         -- earlier message in a thread, handled with the latest one
);

create type public.recon_status as enum (
  'found_in_ledger',
  'received_by_email',
  'amount_mismatch',
  'missing'
);

-- ---------------------------------------------------------------------------
-- Mailboxes
-- ---------------------------------------------------------------------------
create table public.mailboxes (
  id                     uuid primary key default gen_random_uuid(),
  owner_user_id          uuid not null references auth.users (id) on delete cascade,
  provider               public.mail_provider not null,
  email_address          text not null,
  display_name           text,
  -- Vault secret holding provider credentials (refresh token or IMAP password JSON).
  credentials_secret_id  uuid,
  -- Settings (mirror of the ap-inbox-action-report config.md)
  smart_ap_address       text,
  auto_forward_invoices  boolean not null default false,
  lookback_hours         integer not null default 24 check (lookback_hours between 1 and 24 * 31),
  report_recipient       text,           -- defaults to email_address when null
  processed_label        text not null default 'AI-Processed',
  portal_vendors         jsonb not null default '[{"name":"Evergy","portal":"https://www.evergy.com/"}]',
  managed_entities       text[] not null default '{}',  -- properties / entities we manage
  active                 boolean not null default true,
  last_run_at            timestamptz,
  created_at             timestamptz not null default now(),
  unique (owner_user_id, email_address)
);

-- ---------------------------------------------------------------------------
-- Runs
-- ---------------------------------------------------------------------------
create table public.runs (
  id              uuid primary key default gen_random_uuid(),
  mailbox_id      uuid not null references public.mailboxes (id) on delete cascade,
  started_at      timestamptz not null default now(),
  finished_at     timestamptz,
  status          text not null default 'running' check (status in ('running', 'completed', 'failed')),
  emails_seen     integer not null default 0,
  counts          jsonb not null default '{}',
  report_sent_to  text,
  error           text
);
create index runs_mailbox_idx on public.runs (mailbox_id, started_at desc);

-- ---------------------------------------------------------------------------
-- Emails (one row per processed message = one action item)
-- ---------------------------------------------------------------------------
create table public.emails (
  id                   uuid primary key default gen_random_uuid(),
  mailbox_id           uuid not null references public.mailboxes (id) on delete cascade,
  run_id               uuid references public.runs (id) on delete set null,
  provider_message_id  text not null,
  thread_id            text,
  from_address         text,
  from_name            text,
  subject              text,
  received_at          timestamptz,
  attachment_names     text[] not null default '{}',
  category             public.email_category not null,
  confidence           text,
  vendor_name          text,
  invoice_number       text,
  amount               numeric(14, 2),
  currency             text,
  due_date             date,
  summary              text,
  reasoning            text,
  extracted            jsonb not null default '{}',
  flags                text[] not null default '{}',
  status               public.item_status not null,
  action_taken         text,       -- what the app did (e.g. "Forwarded to Smart AP")
  employee_action      text,       -- exactly what the employee must do next
  forwarded_at         timestamptz,
  labeled_at           timestamptz,
  completed_at         timestamptz,
  completed_by         uuid references auth.users (id),
  error                text,
  created_at           timestamptz not null default now(),
  unique (mailbox_id, provider_message_id)
);
create index emails_open_idx on public.emails (mailbox_id, status) where status in ('action_needed', 'flagged');
create index emails_invoice_idx on public.emails (mailbox_id, upper(regexp_replace(invoice_number, '[^A-Za-z0-9]', '', 'g')));

-- ---------------------------------------------------------------------------
-- Statement lines (category D reconciliation results)
-- ---------------------------------------------------------------------------
create table public.statement_lines (
  id              uuid primary key default gen_random_uuid(),
  email_id        uuid not null references public.emails (id) on delete cascade,
  invoice_number  text,
  invoice_date    date,
  amount          numeric(14, 2),
  balance         numeric(14, 2),
  vendor_status   text,              -- open / paid as printed on the statement
  recon_status    public.recon_status not null,
  matched_ref     text,              -- ledger id / email id it matched
  note            text
);
create index statement_lines_email_idx on public.statement_lines (email_id);

-- ---------------------------------------------------------------------------
-- AP ledger: invoices known to the AP system (import from Smart AP / QuickBooks / Yardi CSV).
-- Reconciliation checks here first, then invoices received by email.
-- ---------------------------------------------------------------------------
create table public.ap_ledger (
  id              uuid primary key default gen_random_uuid(),
  owner_user_id   uuid not null references auth.users (id) on delete cascade default auth.uid(),
  vendor_name     text not null,
  invoice_number  text not null,
  invoice_date    date,
  amount          numeric(14, 2),
  status          text,
  source          text not null default 'manual',
  created_at      timestamptz not null default now(),
  unique (owner_user_id, vendor_name, invoice_number)
);
create index ap_ledger_inv_idx on public.ap_ledger (owner_user_id, upper(regexp_replace(invoice_number, '[^A-Za-z0-9]', '', 'g')));

-- ---------------------------------------------------------------------------
-- Row level security: users see only their own mailboxes and the data under them.
-- Edge functions use the service role and bypass RLS.
-- ---------------------------------------------------------------------------
alter table public.mailboxes       enable row level security;
alter table public.runs            enable row level security;
alter table public.emails          enable row level security;
alter table public.statement_lines enable row level security;
alter table public.ap_ledger       enable row level security;

create policy "own mailboxes" on public.mailboxes
  for all to authenticated
  using (owner_user_id = auth.uid())
  with check (owner_user_id = auth.uid());

create policy "own runs" on public.runs
  for select to authenticated
  using (exists (select 1 from public.mailboxes m where m.id = mailbox_id and m.owner_user_id = auth.uid()));

create policy "own emails read" on public.emails
  for select to authenticated
  using (exists (select 1 from public.mailboxes m where m.id = mailbox_id and m.owner_user_id = auth.uid()));

create policy "own statement lines" on public.statement_lines
  for select to authenticated
  using (exists (
    select 1 from public.emails e join public.mailboxes m on m.id = e.mailbox_id
    where e.id = email_id and m.owner_user_id = auth.uid()));

create policy "own ledger" on public.ap_ledger
  for all to authenticated
  using (owner_user_id = auth.uid())
  with check (owner_user_id = auth.uid());

-- Credentials never leave the server. Mailboxes are created by the `connect` edge function;
-- clients may read them (minus the secret id), edit settings, and delete them.
revoke all on public.mailboxes from anon, authenticated;
grant select (id, owner_user_id, provider, email_address, display_name, smart_ap_address,
              auto_forward_invoices, lookback_hours, report_recipient, processed_label,
              portal_vendors, managed_entities, active, last_run_at, created_at)
  on public.mailboxes to authenticated;
grant update (display_name, smart_ap_address, auto_forward_invoices, lookback_hours,
              report_recipient, processed_label, portal_vendors, managed_entities, active)
  on public.mailboxes to authenticated;
grant delete on public.mailboxes to authenticated;

-- Emails, runs and statement lines are written only by edge functions.
revoke insert, update, delete on public.emails, public.runs, public.statement_lines from anon, authenticated;

-- ---------------------------------------------------------------------------
-- Vault helpers (service role only)
-- ---------------------------------------------------------------------------
create or replace function public.mailbox_secret_put(p_secret_id uuid, p_secret text, p_name text)
returns uuid
language plpgsql security definer set search_path = '' as $$
begin
  if p_secret_id is null then
    return vault.create_secret(p_secret, p_name);
  end if;
  perform vault.update_secret(p_secret_id, p_secret);
  return p_secret_id;
end $$;

create or replace function public.mailbox_secret_get(p_secret_id uuid)
returns text
language sql security definer set search_path = '' as $$
  select decrypted_secret from vault.decrypted_secrets where id = p_secret_id
$$;

revoke all on function public.mailbox_secret_put(uuid, text, text) from public, anon, authenticated;
revoke all on function public.mailbox_secret_get(uuid) from public, anon, authenticated;
grant execute on function public.mailbox_secret_put(uuid, text, text) to service_role;
grant execute on function public.mailbox_secret_get(uuid) to service_role;

-- ---------------------------------------------------------------------------
-- Reconciliation lookup: find known invoices by normalized number in the AP ledger
-- and in invoices already received by this mailbox (service role only).
-- ---------------------------------------------------------------------------
create or replace function public.find_invoices(p_owner uuid, p_mailbox uuid, p_numbers text[])
returns table (norm text, source text, ref text, vendor_name text, amount numeric, seen_at timestamptz)
language sql stable security definer set search_path = '' as $$
  select upper(regexp_replace(l.invoice_number, '[^A-Za-z0-9]', '', 'g')), 'ledger', l.id::text,
         l.vendor_name, l.amount, l.created_at
    from public.ap_ledger l
   where l.owner_user_id = p_owner
     and upper(regexp_replace(l.invoice_number, '[^A-Za-z0-9]', '', 'g')) = any (p_numbers)
  union all
  select upper(regexp_replace(e.invoice_number, '[^A-Za-z0-9]', '', 'g')), 'email', e.id::text,
         e.vendor_name, e.amount, e.received_at
    from public.emails e
   where e.mailbox_id = p_mailbox
     and e.category = 'invoice'
     and e.invoice_number is not null
     and upper(regexp_replace(e.invoice_number, '[^A-Za-z0-9]', '', 'g')) = any (p_numbers)
$$;
revoke all on function public.find_invoices(uuid, uuid, text[]) from public, anon, authenticated;
grant execute on function public.find_invoices(uuid, uuid, text[]) to service_role;

create or replace function public.increment_run_seen(p_run_id uuid, p_by integer)
returns void language sql security definer set search_path = '' as $$
  update public.runs set emails_seen = emails_seen + p_by where id = p_run_id
$$;
revoke all on function public.increment_run_seen(uuid, integer) from public, anon, authenticated;
grant execute on function public.increment_run_seen(uuid, integer) to service_role;
