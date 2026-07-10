-- =============================================================================
-- 088_ra_artists_source.sql
-- Tag ra_artists by source so RA + Ticketmaster artists coexist (each pull only
-- clears its own source). Fixes: TM lineup artists never appeared in the artist
-- views because pull-ticketmaster only wrote events, and pull-ra-market's
-- window-delete was wiping ALL artists/events (incl. TM).
-- =============================================================================
begin;
alter table public.ra_artists add column if not exists source text not null default 'ra';
create index if not exists idx_ra_artists_source on public.ra_artists(source);
commit;
-- POST: ra_artists.source ('ra' | 'tm').
