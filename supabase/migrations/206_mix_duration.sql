-- =============================================================================
-- 206_mix_duration.sql
-- How long the episode actually is.
--
-- Every surface that showed an episode's length was summing
-- sc_playlist_tracks.duration_ms — the length of the SOURCE TRACKS. That is not
-- the runtime of the mix, and it is wrong in two different directions at once:
--
--   * A DJ set overlaps and cuts its tracks, so the sum over-reports. SHOW 7
--     sums to 109 minutes of track audio.
--   * Worse, a track added from Beatport (source='dj', the Rekordbox route)
--     stores the length of the PREVIEW CLIP. All 24 of SHOW 6's tracks are that
--     shape, averaging 1.8 minutes and as short as 40 seconds, so the episode
--     rendered as "43 min" — an hour-long mix, under-reported by a third.
--
-- The runtime the audience cares about is the published mix's own duration, and
-- sc-connect's `mix_stats` action already fetches that exact track object from
-- api.soundcloud.com to read its play count. It was throwing `duration` away.
-- Now it stores it here, and every surface reads this instead of summing.
--
-- Nullable on purpose, and the pages render NO minutes when it is null. A blank
-- says "not measured"; a number computed from preview clips says "this mix is 43
-- minutes long", which is a confident false claim. LEARNINGS §23.
--
-- Additive: no view drops a column, the deployed dashboard keeps working, and
-- the value backfills for all 7 live episodes the next time mix stats are pulled.
-- =============================================================================
begin;

alter table public.sc_playlists
  add column if not exists mix_duration_ms int;

comment on column public.sc_playlists.mix_duration_ms is
  'Runtime of the published mix in milliseconds, read from the SoundCloud track by sc-connect mix_stats. NEVER compute this by summing sc_playlist_tracks.duration_ms - those are source-track lengths, and for source=''dj'' tracks they are Beatport preview clips. Null means not yet measured; render no duration rather than a wrong one.';

notify pgrst, 'reload schema';

commit;
