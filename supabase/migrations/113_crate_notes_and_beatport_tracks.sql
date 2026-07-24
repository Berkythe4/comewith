-- =============================================================================
-- 113_crate_notes_and_beatport_tracks.sql
-- Testing a station is where the learning happens, and none of it was capturable:
--   * energy      1-5 rating, Keith's own read of where a track sits in a set
--   * comment     what he noticed while auditioning it (before buying)
--   * sample_url  Beatport's preview clip, so a track that isn't on SoundCloud
--                 can still be auditioned on the site and on the phone
-- Notes are PRIVATE working material: get-station never selects them, so they
-- can't reach the public episode page (it does now serve sample_url, which is
-- Beatport's own public preview).
-- =============================================================================
begin;

alter table public.sc_playlist_tracks add column if not exists energy      int;
alter table public.sc_playlist_tracks add column if not exists comment     text;
alter table public.sc_playlist_tracks add column if not exists sample_url  text;
alter table public.sc_playlist_tracks add column if not exists notes_at    timestamptz;

do $$ begin
  alter table public.sc_playlist_tracks add constraint sc_playlist_tracks_energy_chk
    check (energy is null or energy between 1 and 5);
exception when duplicate_object then null; end $$;

-- A track can now come from Beatport directly (bought/promo, never on SoundCloud).
alter table public.sc_playlist_tracks drop constraint if exists sc_playlist_tracks_source_chk;
alter table public.sc_playlist_tracks add constraint sc_playlist_tracks_source_chk
  check (source = any (array['soundcloud', 'manual', 'rekordbox', 'beatport']));

-- Same for the permanent song memory, so a Beatport-sourced song still logs
-- played/passed like everything else.
alter table public.sc_song_log add column if not exists energy  int;
alter table public.sc_song_log add column if not exists comment text;

commit;
