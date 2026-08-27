-- =============================================================================
-- 206_mix_duration.sql
-- How long the episode actually is.
--
-- Every surface that showed an episode's length was summing
-- sc_playlist_tracks.duration_ms — the length of the SOURCE TRACKS. That is not
-- the runtime of the mix, and measuring it against the real durations shows how
-- far off it drifts:
--
--   station_no:   1    2    3    4    5    6    7
--   real (min):  61   65   64   60   58   43   56
--   summed:      86   98   65   60   59   43  109
--
--   * A DJ set cuts and overlaps its tracks, so summing full SoundCloud tracks
--     OVER-reports — badly. SHOW 7 read 109 minutes for a 56-minute mix.
--   * A track added from Beatport (source='dj', the Rekordbox route) stores the
--     length of the PREVIEW CLIP — all 24 of SHOW 6's are that shape, averaging
--     1.8 min and as short as 40 seconds. Those sums land NEAR the truth (3-6
--     above) purely by coincidence: short clips, and roughly as many of them as
--     a real hour needs. Nothing about that is a measurement, and it breaks the
--     moment an episode mixes the two sources.
--
-- So the old number was not "a bit off". It was a quantity that has no fixed
-- relationship to the runtime and happened to be close on four of seven.
--
-- The runtime the audience cares about is the published mix's own duration, and
-- sc-connect's `mix_stats` action already fetches that exact track object from
-- api.soundcloud.com to read its play count. It was throwing `duration` away.
-- Now it stores it here, and every surface reads this instead of summing.
--
-- Nullable on purpose, and the pages render NO minutes when it is null. A blank
-- says "not measured"; a summed number says "this mix is 109 minutes long" when it
-- runs 56, which is a confident false claim. LEARNINGS §23.
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
