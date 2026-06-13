-- =============================================================================
-- 030_public_events.down.sql — rollback for 030.
--
-- Drops the public-events FEATURE: the view (and its anon/authenticated grant)
-- plus the two new columns.
--
-- It deliberately does NOT restore anon's grant on the events table or the
-- "Public can read non-cancelled events" policy. "Re-lock the DB" keeps the
-- security fix in place; re-opening the pre-030 anon exposure would be a
-- regression.
--
-- If byte-for-byte pre-030 state is ever truly required (NOT recommended — it
-- re-exposes every event column to the world), re-run manually:
--
--   grant all on public.events to anon;
--   create policy "Public can read non-cancelled events" on public.events
--     for select using (status <> 'cancelled' and deleted_at is null);
--
-- =============================================================================
begin;

drop view if exists public.v_public_events;   -- also removes its anon/authenticated grants

alter table public.events drop column if exists ticket_label;
alter table public.events drop column if exists is_public;
-- ticket_url is pre-existing (007_events.sql) — left intact.

commit;
