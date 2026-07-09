-- =============================================================================
-- 081_sc_artist_cache.sql
-- Producer detection + per-artist song cache for the Radio tool. Keyed by the
-- SoundCloud profile URL so it SURVIVES RA re-pulls (pull-ra-market deletes +
-- recreates ra_artists every run — enrichment must not live there). Filled by
-- the sc-enrich edge fn. Admin RLS; grants inherited from 013.
-- Also: ra_artists gains genres[] (from the show they're playing) for filtering.
-- =============================================================================
begin;

create table if not exists public.sc_artist_cache (
  soundcloud text primary key,          -- normalized profile url (lowercase, no trailing slash)
  sc_user_id text,
  username text,
  avatar_url text,
  followers int,
  is_producer boolean default false,    -- has >=1 original short song (not just DJ sets)
  song_count int default 0,
  set_count int default 0,
  songs jsonb default '[]',             -- [{sc_track_id,title,permalink_url,duration_ms,playback_count,created_at,artwork_url}]
  ok boolean default true,              -- false = resolve/tracks failed (dead/renamed profile)
  scanned_at timestamptz not null default now()
);
create index if not exists idx_sc_cache_producer on public.sc_artist_cache(is_producer);

alter table public.sc_artist_cache enable row level security;
drop policy if exists "Admins manage sc_artist_cache" on public.sc_artist_cache;
create policy "Admins manage sc_artist_cache" on public.sc_artist_cache for all using (public.is_admin()) with check (public.is_admin());

alter table public.ra_artists add column if not exists genres text[] default '{}';

commit;
-- POST: producer/song cache (anon-blocked); ra_artists.genres ready (populated
-- on the next RA pull).
