-- =============================================================================
-- 078_ra_market.sql
-- Resident Advisor PUBLIC market data (no auth, no private/ticketing data —
-- see project memory: RA money/guestlist is auth-gated, stays CSV). Powers two
-- internal tools: (1) "RA Market" — scheduling intelligence (best days to throw,
-- venue/genre benchmarks); (2) a weekly radio playlist from upcoming artists.
-- Data is fetched server-side by the pull-ra-market edge function and cached
-- here. Admin-only RLS. Grants inherited from 013 default privileges.
-- =============================================================================
begin;

-- One row per RA event in the pulled window.
create table if not exists public.ra_events (
  ra_id text primary key,
  title text,
  event_date date,
  start_time timestamptz,
  venue_name text,
  area_id int,
  attending int default 0,
  interested_count int default 0,
  is_ticketed boolean default false,
  is_pick boolean default false,            -- RA editorial "pick"
  genres text[] default '{}',
  flyer_url text,
  content_url text,                          -- ra.co/events/<id>
  lineup jsonb default '[]',                 -- [{ra_id,name,soundcloud,follower_count,content_url}]
  fetched_at timestamptz not null default now()
);
create index if not exists idx_ra_events_date on public.ra_events(event_date);

-- Deduped upcoming artists (for the radio station), each with their soonest show.
create table if not exists public.ra_artists (
  ra_id text primary key,
  name text,
  soundcloud text,
  instagram text,
  follower_count int,
  image text,
  content_url text,
  next_event_date date,
  next_event_title text,
  next_venue text,
  fetched_at timestamptz not null default now()
);
create index if not exists idx_ra_artists_next on public.ra_artists(next_event_date);

alter table public.ra_events enable row level security;
alter table public.ra_artists enable row level security;
drop policy if exists "Admins manage ra_events" on public.ra_events;
create policy "Admins manage ra_events" on public.ra_events for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage ra_artists" on public.ra_artists;
create policy "Admins manage ra_artists" on public.ra_artists for all using (public.is_admin()) with check (public.is_admin());

-- Nav: Insights group, after Site Review.
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('ra-market', 'RA Market', 'Insights', 196, true, false, false, '{marketing,full}')
on conflict (key) do nothing;

-- Default RA area code for pulls (8 = New York). Editable in Site Editor → Dashboard settings.
insert into public.site_content (key, value) values ('ops.ra_area_id', '8')
on conflict (key) do nothing;

commit;
-- POST: two admin-only cache tables + 'ra-market' module (master-only until
-- signed off); anon cannot read either (verify with a REST GET → 401).
