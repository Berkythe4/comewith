-- =============================================================================
-- 111_sc_track_count.sql
-- Record what SoundCloud SAYS a profile holds, next to what we could actually
-- read. Claude VonStroke declares track_count = 147; /users/{id}/tracks exposes
-- 27 to an anonymous client and no paging strategy gets further (verified
-- against limit=20/50/200, linked_partitioning, manual offsets and search).
-- Storing the declared number turns a silent shortfall into a visible one.
-- =============================================================================
begin;
alter table public.sc_artist_cache add column if not exists sc_track_count int;
commit;
