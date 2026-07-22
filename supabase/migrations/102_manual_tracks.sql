-- =============================================================================
-- 102_manual_tracks.sql
-- Rekordbox-first station building. SoundCloud is no longer where the mix gets
-- arranged (recording quality) — Keith buys the songs and arranges in Rekordbox,
-- so a station track can now come from somewhere other than the SoundCloud pull:
--
--   • source = 'soundcloud' — added from an artist's ＋add in the Radio panel
--     (sc_track_id is the real SoundCloud numeric id)
--   • source = 'manual'     — typed in by hand (bought / promo / unreleased)
--   • source = 'rekordbox'  — created by the Rekordbox order import for a line
--                             in the export that matched nothing in the station
--
-- sc_track_id stays NOT NULL and stays the key of the sc_song_log memory. Non-
-- SoundCloud tracks get a SYNTHETIC id ('man_<uuid>') minted client-side, so the
-- unique(playlist_id, sc_track_id) dedupe, the played/passed/carried log and the
-- carry-over at finalize all keep working unchanged for bought songs.
--
-- permalink_url is already nullable — a manual track with no public link renders
-- as a text-only row on the episode page (no ▶ / ♡ buttons; radio.html guards).
-- buy_url/label are PRIVATE bookkeeping: get-station does not select them, so
-- they never reach the public page.
-- =============================================================================
begin;

alter table public.sc_playlist_tracks add column if not exists source text not null default 'soundcloud';
alter table public.sc_playlist_tracks drop constraint if exists sc_playlist_tracks_source_chk;
alter table public.sc_playlist_tracks add constraint sc_playlist_tracks_source_chk
  check (source in ('soundcloud', 'manual', 'rekordbox'));
alter table public.sc_playlist_tracks add column if not exists buy_url text;   -- where it was bought (Beatport/Bandcamp) — private
alter table public.sc_playlist_tracks add column if not exists label text;     -- record label — private

-- Song memory carries the same provenance so the 📜 History view can tell a
-- bought song from a SoundCloud pull, and re-adds keep their buy link.
alter table public.sc_song_log add column if not exists source text not null default 'soundcloud';
alter table public.sc_song_log add column if not exists buy_url text;
alter table public.sc_song_log add column if not exists label text;

commit;
-- POST: sc_playlist_tracks + sc_song_log carry source/buy_url/label. No new
-- tables, no policy changes (both tables are already admin-RLS'd + anon-blocked),
-- no grants. Existing rows default to source='soundcloud'.
