-- 125: per-track release date for radio songs. The YouTube episode cards now show
-- each song's genre (already stored in sc_playlist_tracks.genres) AND its release
-- date. Beatport already returns this (publish_date / new_release_date); track-
-- sources fills it in going forward. Nullable text (store the raw date string,
-- e.g. '2023-11-10' or '2023') — the render shows the year and omits when null.
-- Additive only: inherits grants from 013 default privileges; no RLS change.
alter table public.sc_playlist_tracks
  add column if not exists release_date text;

comment on column public.sc_playlist_tracks.release_date is
  'Song release date from the store (Beatport publish_date/new_release_date). Filled by track-sources; only ever set when empty (never overwrites Rekordbox-owned data).';
