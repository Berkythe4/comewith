-- =============================================================================
-- 085_ra_events_source.sql
-- Multi-source upcoming events: tag each ra_events row with its source so RA and
-- Ticketmaster (EDM-only) can coexist in the Market tools. TM events have no
-- RSVP ("attending"); the best-nights DEMAND metric stays RA-only (attending not
-- null), while TM still adds to competition (event count) + venue/genre coverage.
-- =============================================================================
begin;
alter table public.ra_events add column if not exists source text not null default 'ra';
create index if not exists idx_ra_events_source on public.ra_events(source);
commit;
-- POST: ra_events.source ('ra' | 'tm'). pull-ticketmaster upserts source='tm'.
