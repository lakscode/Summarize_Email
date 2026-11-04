-- The bell also tells whoever assigned an email when it is done: completion_seen_at is cleared when an assigned email is
-- marked done (unless the person who assigned it did that themselves) and set when they open the bell. Written only by
-- the edge functions (service role); see inbox/index.ts and process.ts.
alter table public.processed_emails add column if not exists completion_seen_at timestamptz;

-- Tasks already done before this don't count as new.
update public.processed_emails set completion_seen_at = now()
where assigned_by is not null and status = 'done' and completion_seen_at is null;

create index if not exists processed_emails_assigned_by_idx
  on public.processed_emails (assigned_by) where assigned_by is not null;
