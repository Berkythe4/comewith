-- 238: Music Hall of Williamsburg - named by a SHOW 12 tracklist song and 11 shows in
-- the listings, with no venue row, so it sat in the Scene Map's "off this map" tray.
-- Geocoded from its street address (66 N 6th St) via OpenStreetMap.

begin;

insert into public.venues (name, address, lat, lng, source)
select 'Music Hall of Williamsburg', '66 N 6th St, Brooklyn, NY 11249', 40.7191523, -73.9617836, 'crm'
where public.resolve_venue_name('Music Hall of Williamsburg') is null;

update public.ra_events e set venue_id = public.resolve_venue_name(e.venue_name)
 where e.venue_id is null and public.resolve_venue_name(e.venue_name) is not null;

commit;
