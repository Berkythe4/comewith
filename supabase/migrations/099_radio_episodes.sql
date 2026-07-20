-- =============================================================================
-- 099_radio_episodes.sql
-- Come With Radio release pipeline + public episode pages + listener accounts.
--
-- (1) Station lifecycle on sc_playlists: building → testing → live (+ archived).
--     A LIVE station is a published "episode": it gets a slug (pretty public URL
--     radio.html?s=<slug>), a mix (one SoundCloud track = the recorded mix),
--     descriptions (long for the page, short pushed to SoundCloud w/ site link),
--     and published_at. The dashboard's working station = newest non-live row.
--     A partial unique index allows only ONE 'building' row — closes the known
--     raLoadPlaylist race that duplicated empty "Weekly station" rows.
-- (2) radio-mixes storage bucket (public read, admin write) — the uploaded mix
--     audio lives with us; sc-connect streams it to SoundCloud POST /tracks.
-- (3) Listener accounts (role stays 'customer' via handle_new_user): personal
--     playlists + saved tracks + station history. RLS owner-only (or admin).
--     Row-local predicate user_id = auth.uid() lives directly in the policies
--     (INSERT..RETURNING lesson from 097). anon explicitly revoked.
-- Grants otherwise inherited from 013 default privileges. NO anon grants added.
-- =============================================================================
begin;

-- (1) Episode lifecycle -------------------------------------------------------
alter table public.sc_playlists add column if not exists status text not null default 'building'
  check (status in ('building','testing','live','archived'));
alter table public.sc_playlists add column if not exists slug text;
alter table public.sc_playlists add column if not exists mix_file_path text;      -- radio-mixes storage path
alter table public.sc_playlists add column if not exists mix_sc_track_id text;    -- SoundCloud track (the mix)
alter table public.sc_playlists add column if not exists mix_sc_track_url text;
alter table public.sc_playlists add column if not exists mix_youtube_url text;    -- 1.1: manual paste until YT upload API
alter table public.sc_playlists add column if not exists desc_public text;        -- long — shown on the episode page
alter table public.sc_playlists add column if not exists desc_sc text;           -- short — pushed to the SC track, links back to the site
alter table public.sc_playlists add column if not exists published_at timestamptz;
alter table public.sc_playlists add column if not exists cover_url text;

create unique index if not exists uq_sc_playlists_slug on public.sc_playlists (slug) where slug is not null;

-- Exactly one working ('building') station at a time. Archive any extras first
-- (keep the newest row that has tracks; empty dupes came from the known race).
with ranked as (
  select p.id,
         row_number() over (order by (select count(*) from public.sc_playlist_tracks t where t.playlist_id = p.id) desc,
                                     p.created_at desc) rn
  from public.sc_playlists p where p.status = 'building'
)
update public.sc_playlists set status = 'archived'
where id in (select id from ranked where rn > 1);
create unique index if not exists uq_sc_playlists_one_building on public.sc_playlists ((true)) where status = 'building';

-- (1b) Station numbering + song memory ---------------------------------------
-- Every station gets a sequential number (EP 1, 2, …) so past playlists are
-- easy to flip between. sc_song_log is the permanent song memory: every song
-- ever added to a station, whether it was PLAYED (in a live episode) or PASSED
-- (removed while testing), and whether it was auto-CARRIED into the next
-- week's station. Drives the played/passed marks when researching future
-- stations + the "what did I play when" history view.
alter table public.sc_playlists add column if not exists station_no int;
create unique index if not exists uq_sc_playlists_no on public.sc_playlists (station_no) where station_no is not null;
with nums as (
  select id, row_number() over (order by created_at) rn
  from public.sc_playlists where station_no is null
)
update public.sc_playlists p
set station_no = coalesce((select max(station_no) from public.sc_playlists), 0) + nums.rn
from nums where p.id = nums.id;

-- Where a carried-over track came from (the station_no it was cut from).
alter table public.sc_playlist_tracks add column if not exists carried_from int;

create table if not exists public.sc_song_log (
  sc_track_id text primary key,
  title text,
  artist_name text,
  permalink_url text,
  artwork_url text,
  duration_ms int,
  first_added_at timestamptz not null default now(),
  add_count int not null default 1,
  played_playlist_id uuid references public.sc_playlists(id) on delete set null,
  played_station_no int,
  played_at timestamptz,
  passed_playlist_id uuid references public.sc_playlists(id) on delete set null,
  passed_station_no int,
  passed_at timestamptz,
  carried_to uuid references public.sc_playlists(id) on delete set null,  -- set once; stops re-carry loops
  updated_at timestamptz not null default now()
);
alter table public.sc_song_log enable row level security;
drop policy if exists "Admins manage sc_song_log" on public.sc_song_log;
create policy "Admins manage sc_song_log" on public.sc_song_log for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.sc_song_log from anon;

-- (2) Mix storage bucket ------------------------------------------------------
insert into storage.buckets (id, name, public) values ('radio-mixes', 'radio-mixes', true)
on conflict (id) do nothing;
drop policy if exists "Admins insert radio mixes" on storage.objects;
create policy "Admins insert radio mixes" on storage.objects for insert to authenticated
  with check (bucket_id = 'radio-mixes' and public.is_admin());
drop policy if exists "Admins update radio mixes" on storage.objects;
create policy "Admins update radio mixes" on storage.objects for update to authenticated
  using (bucket_id = 'radio-mixes' and public.is_admin());
drop policy if exists "Admins delete radio mixes" on storage.objects;
create policy "Admins delete radio mixes" on storage.objects for delete to authenticated
  using (bucket_id = 'radio-mixes' and public.is_admin());
drop policy if exists "Admins read radio mixes" on storage.objects;
create policy "Admins read radio mixes" on storage.objects for select to authenticated
  using (bucket_id = 'radio-mixes' and public.is_admin());

-- (3) Listener accounts -------------------------------------------------------
create table if not exists public.listener_playlists (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  name text not null default 'My tracks',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists idx_listener_playlists_user on public.listener_playlists (user_id);

create table if not exists public.listener_playlist_tracks (
  id uuid primary key default gen_random_uuid(),
  playlist_id uuid not null references public.listener_playlists(id) on delete cascade,
  title text,
  artist_name text,
  permalink_url text not null,           -- SoundCloud song link (also the dedupe key)
  artwork_url text,
  bpm int,
  song_key text,
  camelot text,
  station_slug text,                     -- which episode it was saved from
  station_name text,
  added_at timestamptz not null default now(),
  unique (playlist_id, permalink_url)
);
create index if not exists idx_listener_pl_tracks_pl on public.listener_playlist_tracks (playlist_id);

create table if not exists public.listener_station_history (
  user_id uuid not null references auth.users(id) on delete cascade,
  station_slug text not null,
  station_name text,
  first_seen_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  visits int not null default 1,
  primary key (user_id, station_slug)
);

alter table public.listener_playlists enable row level security;
alter table public.listener_playlist_tracks enable row level security;
alter table public.listener_station_history enable row level security;

drop policy if exists "Own listener playlists" on public.listener_playlists;
create policy "Own listener playlists" on public.listener_playlists for all
  using (user_id = auth.uid() or public.is_admin())
  with check (user_id = auth.uid());

drop policy if exists "Own listener playlist tracks" on public.listener_playlist_tracks;
create policy "Own listener playlist tracks" on public.listener_playlist_tracks for all
  using (exists (select 1 from public.listener_playlists p
                 where p.id = playlist_id and (p.user_id = auth.uid() or public.is_admin())))
  with check (exists (select 1 from public.listener_playlists p
                      where p.id = playlist_id and p.user_id = auth.uid()));

drop policy if exists "Own listener history" on public.listener_station_history;
create policy "Own listener history" on public.listener_station_history for all
  using (user_id = auth.uid() or public.is_admin())
  with check (user_id = auth.uid());

revoke all on public.listener_playlists from anon;
revoke all on public.listener_playlist_tracks from anon;
revoke all on public.listener_station_history from anon;

commit;
-- POST: sc_playlists carries the episode lifecycle (status/slug/mix/descriptions)
-- + sequential station_no; one 'building' row enforced; sc_song_log = permanent
-- played/passed/carried song memory (admin-only); radio-mixes bucket (public
-- read, admin write); three listener tables, owner-RLS'd, anon-revoked. Public
-- reads stay function-only via get-station (service role) — nothing granted to anon.
