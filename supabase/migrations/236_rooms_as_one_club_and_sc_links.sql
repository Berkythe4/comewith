-- 236: rooms count as one club, real pins, Elements located, and SoundCloud links
-- for tracklist songs that came in from Rekordbox (Keith, 2026-11-05).
--
-- 1. ROOMS -> BUILDING. Keith: "count them as one club". A named room inside a
--    building (Elsewhere - Rooftop, Ruins at Knockdown Center, The Sultan Room
--    Rooftop, Superior Ingredients' Roof / Main Room, 99 Scott - Main Room) now
--    resolves to the building. ra_events.venue_name keeps the room the feed named,
--    so the room is still readable and every merge is reversible.
--    NOT merged: BASEMENT. It shares Knockdown Center's address but is its own club
--    with its own name and lineups - a separate business, not a room.
-- 2. Brooklyn Pacha IS Pacha New York (Keith). Pacha's pin was the old Hell's Kitchen
--    site; RA gives 140 Stewart Ave, Brooklyn.
-- 3. Pins checked against RA's own venue address (scripts in the 2026-11-05 session):
--    six radio clubs sat 5-13 km from their address and are moved. Gabriela flagged
--    too, but there the GEOCODER was wrong (street-level hit), so it is left alone.
-- 4. Elements Festival = Pocono Raceway, Long Pond PA (Keith). ONE 77 = 177 2nd Ave,
--    Brooklyn (DICE's event page).
-- 5. sc_playlist_tracks.sc_link_*: a SoundCloud track found for a song whose own id is
--    synthetic (Rekordbox cue 'cue_…' / hand-added 'man_…'). sc_track_id stays the
--    song's identity - sc_song_log, dedupe and carry-over key on it - so the link is
--    a separate field, written only by scripts/link_radio_tracks_soundcloud.py.

begin;

create function pg_temp.merge_venue(p_dup text, p_keep text) returns void language plpgsql as $$
declare d public.venues; k public.venues;
begin
  if (select count(*) from public.venues where name = p_dup and deleted_at is null) <> 1 then raise exception 'merge: no single live venue named %', p_dup; end if;
  if (select count(*) from public.venues where name = p_keep and deleted_at is null) <> 1 then raise exception 'merge: no single live venue named %', p_keep; end if;
  select * into d from public.venues where name = p_dup and deleted_at is null;
  select * into k from public.venues where name = p_keep and deleted_at is null;
  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
  values (d.name_norm, btrim(d.name), k.id, 'linked', 'manual', '236: room of ' || k.name)
  on conflict (alias_norm) do update set venue_id = excluded.venue_id, status = 'linked', note = excluded.note;
  update public.venue_aliases set venue_id = k.id where venue_id = d.id;
  update public.ra_events     set venue_id = k.id where venue_id = d.id;
  update public.events        set venue_id = k.id where venue_id = d.id;
  update public.venue_contacts set venue_id = k.id where venue_id = d.id;
  update public.venues set deleted_at = now(), updated_at = now(),
         notes = concat_ws(E'\n', notes, '236: room folded into ' || k.name || ' (' || k.id || ')')
   where id = d.id;
end $$;

create function pg_temp.alias(p_raw text, p_keep text) returns void language plpgsql as $$
declare k uuid; n text := public.normalize_venue_name(p_raw);
begin
  if (select count(*) from public.venues where name = p_keep and deleted_at is null) <> 1 then raise exception 'alias: no single live venue named %', p_keep; end if;
  select id into k from public.venues where name = p_keep and deleted_at is null;
  if exists (select 1 from public.venues where name_norm = n and deleted_at is null and id <> k) then raise exception 'alias: % is itself a venue row - merge it instead', p_raw; end if;
  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
  values (n, btrim(p_raw), k, 'linked', 'manual', '236: room / spelling of ' || p_keep)
  on conflict (alias_norm) do update set venue_id = excluded.venue_id, status = 'linked', note = excluded.note;
  update public.ra_events set venue_id = k where public.normalize_venue_name(venue_name) = n and venue_id is distinct from k;
end $$;

-- 1. Rooms with their own venue row.
select pg_temp.merge_venue('Elsewhere - Rooftop', 'Elsewhere');
select pg_temp.merge_venue('H0l0 Yard', 'H0L0');
select pg_temp.merge_venue('Creekside - Under The K Bridge', 'Under the K Bridge');
select pg_temp.merge_venue('Superior Ingredients (The Roof + Room)', 'Superior Ingredients');
select pg_temp.merge_venue('The Sultan Room Rooftop', 'The Sultan Room');
select pg_temp.merge_venue('The Crown (Rooftop)', 'The Crown');
select pg_temp.merge_venue('Rooftop at Arlo Williamsburg ', 'Arlo Williamsburg');
select pg_temp.merge_venue('Althea''s Rooftop', 'Althea''s Hideaway NYC');
-- Rooms that only ever appeared as a spelling.
select pg_temp.alias('Chatroom at Elsewhere', 'Elsewhere');
select pg_temp.alias('Elsewhere - The Hall', 'Elsewhere');
select pg_temp.alias('Elsewhere - Zone One', 'Elsewhere');
select pg_temp.alias('Ruins at Knockdown Center', 'Knockdown Center');
select pg_temp.alias('Superior Ingredients (Main Room)', 'Superior Ingredients');
select pg_temp.alias('The Roof at Superior Ingredients', 'Superior Ingredients');
select pg_temp.alias('99 Scott - Main Room', '99 Scott Ave');
select pg_temp.alias('99 Scott - Courtyard', '99 Scott Ave');
select pg_temp.alias('Sunset Flex - Under The K Bridge', 'Under the K Bridge');
select pg_temp.alias('Mirror Bar At Arlo Williamsburg', 'Arlo Williamsburg');
select pg_temp.alias('Arlo Williamsburg (Formerly The Williamsburg Hotel)', 'Arlo Williamsburg');
select pg_temp.alias('Brooklyn Army Terminal: Pier 4', 'Brooklyn Army Terminal');
-- 2.
select pg_temp.alias('Brooklyn Pacha', 'Pacha New York');

-- 3. Pins, from RA's venue address.
update public.venues v set lat = x.lat, lng = x.lng, address = coalesce(nullif(v.address, ''), x.address), updated_at = now()
from (values
  ('Pacha New York',        40.7106189, -73.9261811, '140 Stewart Ave, Brooklyn, NY 11237'),
  ('Refuge',                40.7117822, -73.9351694, '366 Ten Eyck St, Brooklyn, NY 11206'),
  ('The Chocolate Factory', 40.7098872, -73.9224221, '70 Scott Ave, Brooklyn, NY 11237'),
  ('Outer Heaven',          40.7219976, -73.9922737, '191 Chrystie St, New York, NY 10002'),
  ('Superior Ingredients',  40.7213,    -73.9590,    '74 Wythe Ave, Brooklyn, NY 11249'),
  ('ErF Studios',           40.7069298, -73.9334563, '1 Ingraham St, Brooklyn, NY 11206'),
  ('Brooklyn Storehouse',   40.6993266, -73.9755562, 'Brooklyn Navy Yard, Brooklyn, NY')
) as x(name, lat, lng, address)
where v.name = x.name and v.deleted_at is null;
-- Pacha's old pin was the Hell's Kitchen club; its address must not survive the move.
update public.venues set address = '140 Stewart Ave, Brooklyn, NY 11237' where name = 'Pacha New York' and deleted_at is null;

-- 4. New rooms.
insert into public.venues (name, address, area, lat, lng, source)
select x.name, x.address, x.area, x.lat, x.lng, 'crm'
from (values
  ('Elements Festival', 'Pocono Raceway, 1234 Long Pond Rd, Long Pond, PA 18334', 'Pennsylvania (festival)', 41.0605146, -75.5117578),
  ('ONE 77',            '177 2nd Ave, Brooklyn, NY 11215',                         null,                      40.669632,  -73.99561)
) as x(name, address, area, lat, lng)
where public.resolve_venue_name(x.name) is null;
update public.ra_events e set venue_id = public.resolve_venue_name(e.venue_name)
 where e.venue_id is null and public.resolve_venue_name(e.venue_name) is not null;

-- 5. SoundCloud link for synthetic-id songs.
alter table public.sc_playlist_tracks
  add column if not exists sc_link_id         text,
  add column if not exists sc_link_url        text,
  add column if not exists sc_link_how        text,
  add column if not exists sc_link_checked_at timestamptz;
alter table public.sc_playlist_tracks drop constraint if exists sc_playlist_tracks_sc_link_chk;
alter table public.sc_playlist_tracks add constraint sc_playlist_tracks_sc_link_chk
  check (sc_link_id is null or sc_link_id ~ '^\d+$');

commit;
