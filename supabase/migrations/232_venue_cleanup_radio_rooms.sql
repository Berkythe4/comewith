-- 232: venue clean-up for the radio Scene Map (2026-11-05, Keith: "combine obvious
-- duplicates"; the tracklists' own show_venue is now the map's primary source).
--
-- THE RULE USED: merge only when two names are the SAME room spelled differently
-- (a city suffix, a "TBA - " prefix, an address appended, an apostrophe). A named
-- ROOM inside a building (Elsewhere - Rooftop, Ruins at Knockdown Center, The Sultan
-- Room Rooftop, Superior Ingredients (Main Room), 99 Scott - Main Room) is NOT merged
-- here - that is a decision about what a "club" is on the map, left to Keith.
-- LEARNINGS §63: a wrong merge rewrites history, so every merge below is a spelling,
-- and every one is reversible (aliases + soft delete; ra_events.venue_name untouched).
--
-- Mechanics: venue_aliases rows are written directly (link_venue_alias() requires an
-- admin session, which a migration does not have) and ra_events re-pointed in the
-- same statement set, exactly as link_venue_alias() does it.

begin;

-- Merge venue row `dup` into `keep`: alias dup's name, re-point everything, soft-delete.
create function pg_temp.merge_venue(p_dup text, p_keep text) returns void language plpgsql as $$
declare d public.venues; k public.venues;
begin
  select * into d from public.venues where name = p_dup and deleted_at is null;
  if not found or (select count(*) from public.venues where name = p_dup and deleted_at is null) <> 1 then raise exception 'merge: no single live venue named %', p_dup; end if;
  select * into k from public.venues where name = p_keep and deleted_at is null;
  if not found or (select count(*) from public.venues where name = p_keep and deleted_at is null) <> 1 then raise exception 'merge: no single live venue named %', p_keep; end if;
  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
  values (d.name_norm, btrim(d.name), k.id, 'linked', 'manual', '232: duplicate venue row merged into ' || k.name)
  on conflict (alias_norm) do update set venue_id = excluded.venue_id, status = 'linked', note = excluded.note;
  update public.venue_aliases set venue_id = k.id where venue_id = d.id;
  update public.ra_events     set venue_id = k.id where venue_id = d.id;
  update public.events        set venue_id = k.id where venue_id = d.id;
  update public.venue_contacts set venue_id = k.id where venue_id = d.id;
  update public.venues set deleted_at = now(), updated_at = now(),
         notes = concat_ws(E'\n', notes, '232: merged into ' || k.name || ' (' || k.id || ')')
   where id = d.id;
end $$;

-- A spelling with no venue row of its own -> an existing venue.
create function pg_temp.alias(p_raw text, p_keep text) returns void language plpgsql as $$
declare k uuid; n text := public.normalize_venue_name(p_raw);
begin
  select id into k from public.venues where name = p_keep and deleted_at is null;
  if k is null or (select count(*) from public.venues where name = p_keep and deleted_at is null) <> 1 then raise exception 'alias: no single live venue named %', p_keep; end if;
  if exists (select 1 from public.venues where name_norm = n and deleted_at is null and id <> k) then raise exception 'alias: % is itself a venue row - merge it instead', p_raw; end if;
  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
  values (n, btrim(p_raw), k, 'linked', 'manual', '232: spelling of ' || p_keep)
  on conflict (alias_norm) do update set venue_id = excluded.venue_id, status = 'linked', note = excluded.note;
  update public.ra_events set venue_id = k where public.normalize_venue_name(venue_name) = n and venue_id is distinct from k;
end $$;

-- 1. Duplicate venue ROWS (same room, two records).
select pg_temp.merge_venue('SILO Brooklyn ', 'SILO');
select pg_temp.merge_venue('Elsewhere, Brooklyn', 'Elsewhere');
select pg_temp.merge_venue('TBA - Reforesters Laboratory', 'Reforesters Laboratory');
select pg_temp.merge_venue('Randalls Island', 'Randall''s Island');
select pg_temp.merge_venue('Circle Line Cruises, Pier 83', 'Circle Line Cruises');
select pg_temp.merge_venue('314 Scholes St, Brooklyn, NY 11206, USA', '314 Scholes');
select pg_temp.merge_venue('TBA - Brooklyn Brewery (79 N 11th St, Brooklyn, NY 11249)', 'Brooklyn Brewery');
select pg_temp.merge_venue('TBA - Mr.Purple 180 Orchard St, New York, NY 10002 ', 'Mr. Purple');

-- 2. Spellings from the review queue (v_venue_name_review) that are the same room.
select pg_temp.alias('314 Scholes St', '314 Scholes');
select pg_temp.alias('314 Scholes Street', '314 Scholes');
select pg_temp.alias('Bk Backyard Bar', 'BK Backyard');
select pg_temp.alias('Hotel 50 Bowery NY', 'Hotel 50 Bowery');
select pg_temp.alias('Magic Hour Rooftop Bar', 'Magic Hour Rooftop');
select pg_temp.alias('315 Meserole St.', '315 Meserole');
select pg_temp.alias('TBA - Lella Alimentari', 'Lella Alimentari');
select pg_temp.alias('Under the ''K'' Bridge Park', 'Under the K Bridge');
select pg_temp.alias('TBA - Brooklyn Brewery', 'Brooklyn Brewery');
select pg_temp.alias('Substance Skate Park', 'Substance Skatepark');
select pg_temp.alias('Randall''s Island Park', 'Randall''s Island');
select pg_temp.alias('Racket NYC', 'Racket NY');
select pg_temp.alias('Boyfriend Coop', 'Boyfriend co-op');
select pg_temp.alias('Nightclub 101', 'Night Club 101');
select pg_temp.alias('Circle Line Sightseeing Cruises', 'Circle Line Cruises');
select pg_temp.alias('Circle Line Boat', 'Circle Line Cruises');
select pg_temp.alias('Baby''s All Right', 'TBA - Babys All Right');
-- ...and give that room its real name. The old spelling keeps resolving via an alias.
insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
select name_norm, btrim(name), id, 'linked', 'manual', '232: old name of this row'
  from public.venues where name = 'TBA - Babys All Right' and deleted_at is null
on conflict (alias_norm) do nothing;
update public.venues set name = 'Baby''s All Right', updated_at = now() where name = 'TBA - Babys All Right' and deleted_at is null;

-- 3. Wrong locations. SILO sat at 40.56,-73.89 (near Rockaway Inlet); RA gives
--    90 Scott Ave, East Williamsburg. H0l0 Yard sat at Hudson Yards; it is H0L0's
--    outdoor space in Ridgewood (kept as its own room, moved to H0L0's point).
update public.venues set lat = 40.7105087, lng = -73.922873, address = coalesce(address, '90 Scott Ave, Brooklyn, NY 11237'), updated_at = now()
 where name = 'SILO' and deleted_at is null;
update public.venues y set lat = h.lat, lng = h.lng, updated_at = now()
  from public.venues h where y.name = 'H0l0 Yard' and y.deleted_at is null and h.name = 'H0L0' and h.deleted_at is null;

-- 4. Rooms our tracklists name that had no venue row (geocoded from their street
--    address via OpenStreetMap; SummerStage placed at Rumsey Playfield by hand).
--    Not created: Elements Festival (location unknown), ONE 77 and Brooklyn Pacha
--    (the geocoder found the wrong place) - they stay listed as "not on the map".
insert into public.venues (name, address, lat, lng, source)
select x.name, x.address, x.lat, x.lng, 'crm'
from (values
  ('Westlight Rooftop at The William Vale', '111 N 12th St, Brooklyn, NY', 40.7224118, -73.9569215),
  ('Flushing Meadows Corona Park', 'Flushing Meadows Corona Park, Queens, NY', 40.7405689, -73.8428304),
  ('Brooklyn Paramount', '385 Flatbush Ave Ext, Brooklyn, NY', 40.6894873, -73.9810419),
  ('Central Park SummerStage', 'Rumsey Playfield, Central Park, New York, NY', 40.7712, -73.9717),
  ('Capitale', '130 Bowery, New York, NY', 40.7189317, -73.9952628),
  ('Pianos', '158 Ludlow St, New York, NY', 40.7210032, -73.9876202),
  ('Cipriani Wall Street', '55 Wall St, New York, NY', 40.7059032, -74.0093449),
  ('Marquee Skydeck at Edge Hudson Yards', '30 Hudson Yards, New York, NY', 40.753983, -74.0006028)
) as x(name, address, lat, lng)
where not exists (select 1 from public.venues v where v.name_norm = public.normalize_venue_name(x.name) and v.deleted_at is null)
  and public.resolve_venue_name(x.name) is null;

-- Point the shows already stored under those names at the new rows.
update public.ra_events e set venue_id = public.resolve_venue_name(e.venue_name)
 where e.venue_id is null and public.resolve_venue_name(e.venue_name) is not null;

commit;
