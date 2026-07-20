-- =============================================================================
-- 101_track_genres.sql
-- Genre chips on station tracks (snapshotted from ra_artists.genres at add time,
-- refreshable via the dashboard "↻ Show info" action, shown on the public
-- episode page + CSV). Also the columns ride along to listener saves later.
-- =============================================================================
begin;
alter table public.sc_playlist_tracks add column if not exists genres text[];
commit;
-- POST: sc_playlist_tracks.genres.
