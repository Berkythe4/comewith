-- 132: allow source='dj' on tracks so a DJ's picks (added via the scoped dj.html
-- workspace) are distinguishable + only-their-own-removable. Extends the existing
-- source allow-list.
alter table public.sc_playlist_tracks drop constraint if exists sc_playlist_tracks_source_chk;
alter table public.sc_playlist_tracks add constraint sc_playlist_tracks_source_chk
  check (source = any (array['soundcloud','manual','rekordbox','beatport','dj']));
