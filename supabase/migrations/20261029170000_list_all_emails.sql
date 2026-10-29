-- Back to listing every email: mail set aside as not relevant returns to "other" / no action, as before
-- 20261029160000_not_relevant_emails. The extra allowed category and status values stay; nothing uses them now.
update public.processed_emails
set category = 'other', status = 'no_action', updated_at = now()
where status = 'ignored' or category = 'not_relevant';
