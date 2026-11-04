-- How an email / a mailbox was last processed, next to who: 'manual' (someone pressed Process inbox), 'schedule' (the
-- scheduler, which runs as the mailbox owner, so processed_by alone would read as if the owner did it) or 'user' (someone
-- marked it done / reopened it or changed its category), as email_processing_log.trigger already records. Written only by
-- the edge functions (service role).
alter table public.processed_emails
  add column if not exists processed_trigger text check (processed_trigger in ('manual', 'schedule', 'user'));
alter table public.mailboxes
  add column if not exists last_processed_trigger text check (last_processed_trigger in ('manual', 'schedule'));

grant select (last_processed_trigger) on public.mailboxes to authenticated;

-- Existing rows: from their latest processing history row.
update public.processed_emails e
   set processed_trigger = l.trigger
  from (select distinct on (processed_email_id) processed_email_id, trigger
          from public.email_processing_log where processed_email_id is not null
         order by processed_email_id, processed_at desc, id desc) l
 where l.processed_email_id = e.id and e.processed_trigger is null;

update public.mailboxes m
   set last_processed_trigger = l.trigger
  from (select distinct on (mailbox_id) mailbox_id, trigger
          from public.email_processing_log where mailbox_id is not null and trigger in ('manual', 'schedule')
         order by mailbox_id, processed_at desc, id desc) l
 where l.mailbox_id = m.id and m.last_processed_trigger is null;
