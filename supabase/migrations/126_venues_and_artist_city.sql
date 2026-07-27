-- 126: two foundations for the Best Nights rework.
--
-- (a) ra_artists.city + is_partner. City comes from any easily available source
--     (SoundCloud profile city first) so we can tag NYC-LOCAL artists and filter
--     to them. is_partner flags artists we manually add for a partnership.
-- (b) a first-class `venues` table. The future public heat-map needs venues as
--     real entities (name, area, lat/lng to place a pin, links). Standing it up
--     now — even sparsely filled — means the heat-map doesn't force a rebuild.
--
-- Additive only. ra_artists is admin-managed market cache; venues is admin-RLS'd.

alter table public.ra_artists
  add column if not exists city text,
  add column if not exists is_partner boolean not null default false;

create table if not exists public.venues (
  id           uuid primary key default gen_random_uuid(),
  name         text not null,
  area         text,                       -- neighborhood / borough
  city         text default 'New York',
  capacity     integer,
  lat          double precision,           -- for the heat-map pin (geocode later)
  lng          double precision,
  website      text,
  instagram    text,
  ra_url       text,
  ticket_url   text,
  genres       text[],
  is_partner   boolean not null default false,
  notes        text,
  source       text not null default 'manual',
  ra_id        text,                        -- link back to an ra_events venue if matched
  created_at   timestamptz not null default now(),
  created_by   uuid
);
-- one row per venue name (case-insensitive) so re-adding updates instead of dupes
create unique index if not exists venues_name_key on public.venues (lower(name));

alter table public.venues enable row level security;
drop policy if exists venues_admin on public.venues;
create policy venues_admin on public.venues
  for all using (public.is_admin()) with check (public.is_admin());
