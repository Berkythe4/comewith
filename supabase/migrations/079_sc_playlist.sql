-- =============================================================================
-- 079_sc_playlist.sql
-- In-app SoundCloud playlist builder for the Radio feature. Tracks are pulled by
-- the sc-tracks edge fn (SoundCloud's internal read API — songs only, sets
-- filtered by duration). Playlists live here so they persist + can later be
-- pushed to a real SoundCloud playlist via OAuth (needs Artist Pro app — phase 2).
-- Admin RLS. Grants inherited from 013 default privileges.
-- =============================================================================
begin;

create table if not exists public.sc_playlists (
  id uuid primary key default gen_random_uuid(),
  name text not null default 'Weekly station',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.sc_playlist_tracks (
  id uuid primary key default gen_random_uuid(),
  playlist_id uuid not null references public.sc_playlists(id) on delete cascade,
  sc_track_id text not null,             -- SoundCloud numeric track id
  title text,
  artist_name text,
  permalink_url text,                    -- soundcloud.com/... (for embed + export)
  duration_ms int,
  playback_count int,
  artwork_url text,
  sort int not null default 100,
  added_at timestamptz not null default now(),
  unique (playlist_id, sc_track_id)      -- no dupes in a playlist
);
create index if not exists idx_sc_pl_tracks_pl on public.sc_playlist_tracks(playlist_id, sort);

alter table public.sc_playlists enable row level security;
alter table public.sc_playlist_tracks enable row level security;
drop policy if exists "Admins manage sc_playlists" on public.sc_playlists;
create policy "Admins manage sc_playlists" on public.sc_playlists for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage sc_playlist_tracks" on public.sc_playlist_tracks;
create policy "Admins manage sc_playlist_tracks" on public.sc_playlist_tracks for all using (public.is_admin()) with check (public.is_admin());

-- Cache the currently-working SoundCloud client_id (rotates; sc-tracks refreshes on 401).
insert into public.site_content (key, value) values ('ops.sc_client_id', '')
on conflict (key) do nothing;

commit;
-- POST: two admin-only playlist tables (anon-blocked); a starter client_id cache key.
