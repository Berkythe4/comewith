-- =============================================================================
-- 082_tracklist_show_info.sql
-- Publishable tracklist: capture the SHOW each station track's artist is playing
-- (date / venue / price / RA link) so the exported/shared tracklist reads like a
-- gig guide. RA exposes event.cost publicly (free text: "$20+", "$15-30", "30").
-- =============================================================================
begin;

-- RA cache: the artist's soonest upcoming show cost + link (populated by pull-ra-market).
alter table public.ra_artists add column if not exists next_cost text;
alter table public.ra_artists add column if not exists next_event_url text;

-- Station tracks remember the show context at add-time (RA data churns on re-pull).
alter table public.sc_playlist_tracks add column if not exists show_date date;
alter table public.sc_playlist_tracks add column if not exists show_venue text;
alter table public.sc_playlist_tracks add column if not exists show_cost text;
alter table public.sc_playlist_tracks add column if not exists show_url text;

commit;
-- POST: tracks can carry show date/venue/price/link; ra_artists gains cost+url
-- (filled on the next RA pull).
