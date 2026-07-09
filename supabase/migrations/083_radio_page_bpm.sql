-- =============================================================================
-- 083_radio_page_bpm.sql
-- (1) Public "Come With Radio" page scaffolding — BUILT but not shown: each
--     station gets a public_token + published flag (default FALSE). The page +
--     get-station edge fn only serve a station when published, and nothing links
--     to it, so it stays private until Keith flips it on.
-- (2) BPM + musical key + Camelot on station tracks (filled later by a GetSongBPM
--     enrichment once an API key is added). Helps harmonic-mix the tracklist.
-- =============================================================================
begin;

alter table public.sc_playlists add column if not exists published boolean not null default false;
alter table public.sc_playlists add column if not exists public_token uuid not null default gen_random_uuid();
alter table public.sc_playlists add column if not exists note text;

alter table public.sc_playlist_tracks add column if not exists bpm int;
alter table public.sc_playlist_tracks add column if not exists song_key text;      -- musical, e.g. "Am", "F#"
alter table public.sc_playlist_tracks add column if not exists camelot text;       -- e.g. "8A"

commit;
-- POST: stations have a (default-unpublished) public_token; tracks can hold
-- bpm/key/camelot. Public page serves nothing until published=true.
