-- =============================================================================
-- 105_episode_mix_by.sql
-- Who DJ'd the episode. Each Come With Radio episode is a recorded mix by a
-- specific DJ; capture their name so it shows on the public episode page and in
-- the SoundCloud description ("Mixed by …"). Free text — usually Keith, but
-- guest mixes happen.
-- =============================================================================
begin;
alter table public.sc_playlists add column if not exists mix_by text;
commit;
-- POST: sc_playlists.mix_by (nullable free text).
