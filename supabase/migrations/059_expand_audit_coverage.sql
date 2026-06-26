-- =============================================================================
-- 059_expand_audit_coverage.sql
-- The activity log (Users → Activity) reads audit_log, but only 7 tables had
-- audit triggers, so "everyone's activity" missed most actions. Extend the same
-- audit_trigger_function() to the rest of the user-action tables so the feed is
-- comprehensive. audit_log stays master-only read (no new exposure). High-volume
-- / low-signal tables (conversation_messages, social_post_notes, metric_snapshots,
-- guest_event_attendance) are intentionally left out.
-- =============================================================================
begin;
do $$
declare t text;
begin
  foreach t in array array[
    'actors','events','venues','sponsorships','ticketing','third_party_donations',
    'inquiries','social_posts','conversations','equipment_inventory'
  ] loop
    execute format('drop trigger if exists audit_%1$s on public.%1$s', t);
    execute format('create trigger audit_%1$s after insert or update or delete on public.%1$s for each row execute function public.audit_trigger_function()', t);
  end loop;
end $$;
commit;
-- POST: inserts/updates/deletes on these tables now land in audit_log with the
-- acting user (auth.uid) + email, visible in the Users → Activity log.
