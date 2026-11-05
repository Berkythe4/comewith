-- 233: v_radio_track_shows - every AIRED radio song with the show it was picked for.
--
-- The tracklists are the Scene Map's primary source (Keith, 2026-11-05). Every song
-- on a station already records the show that put it there: sc_playlist_tracks
-- show_date / show_venue / show_url / show_cost (253 of the 256 aired songs carry a
-- venue and a date). Matching lineup names against the whole listings history found
-- 189 of 256 and could only ever guess which night; the tracklist SAYS which night.
--
-- Per song this resolves:
--   venue_id  the room - from the matched event when there is one, else the
--             tracklist's own show_venue through resolve_venue_name() (aliases
--             first, then a venue whose name normalises the same). Never fuzzy.
--   event     the party, tried in order, first hit wins ("how"):
--             'url'     the song's show_url IS the event (RA id, or the same DICE/TM URL)
--             'lineup'  same date, an artist credited on the song is in the lineup,
--                       and the same room (resolved id or normalised name)
--             'room'    same date, same resolved room, and it is the ONLY show there
--                       that day - one candidate, so not a guess
--             null      none of the above; the room still stands on its own.
-- "Aired" = a station that is live or archived. A building station has not aired.
-- The artist split is the dashboard's RADIO_SPLIT_RE; change one, change the other.

begin;

create or replace view public.v_radio_track_shows
with (security_invoker = true) as
with tr as (
  select t.id, t.playlist_id, p.station_no, p.status as station_status, t.sort,
         t.sc_track_id, t.title, t.artist_name, t.permalink_url, t.artwork_url, t.duration_ms,
         t.genres, t.label, t.bpm, t.camelot,
         t.show_date, nullif(btrim(t.show_venue), '') as show_venue,
         nullif(btrim(t.show_url), '') as show_url, t.show_cost,
         public.resolve_venue_name(nullif(btrim(t.show_venue), '')) as show_venue_id,
         array(select btrim(x)
                 from regexp_split_to_table(lower(coalesce(t.artist_name, '')),
                      ',|&|\mfeat\.?|\mft\.?|\mx\M|\mvs\.?') x
                where btrim(x) <> '') as parts
  from public.sc_playlist_tracks t
  join public.sc_playlists p on p.id = t.playlist_id
  where p.status in ('live', 'archived')
)
select
  tr.id, tr.playlist_id, tr.station_no, tr.sort, tr.sc_track_id, tr.title, tr.artist_name,
  tr.permalink_url, tr.artwork_url, tr.duration_ms, tr.genres, tr.label, tr.bpm, tr.camelot,
  tr.show_date, tr.show_venue, tr.show_url, tr.show_cost,
  coalesce(ev.venue_id, tr.show_venue_id)          as venue_id,
  v.name                                           as venue_name,
  ev.ra_id                                         as event_ra_id,
  nullif(btrim(ev.title), '')                      as event_title,
  ev.content_url                                   as event_url,
  ev.source                                        as event_source,
  ev.attending                                     as event_attending,
  ev.how                                           as event_match
from tr
left join lateral (
  select c.* from (
    select e.ra_id, e.title, e.venue_id, e.content_url, e.source, e.attending, 'url'::text as how, 1 as pri
      from public.ra_events e
     where tr.show_url is not null
       and (e.content_url = split_part(tr.show_url, '?', 1)
            or (tr.show_url ~ 'ra\.co/events/\d+' and e.ra_id = substring(tr.show_url from 'ra\.co/events/(\d+)')))
    union all
    select e.ra_id, e.title, e.venue_id, e.content_url, e.source, e.attending, 'lineup', 2
      from public.ra_events e
     where tr.show_date is not null and e.event_date = tr.show_date
       and exists (select 1 from jsonb_array_elements(coalesce(e.lineup, '[]'::jsonb)) l
                    where lower(btrim(l->>'name')) = any(tr.parts))
       and (tr.show_venue is null or e.venue_id = tr.show_venue_id
            or public.normalize_venue_name(e.venue_name) = public.normalize_venue_name(tr.show_venue))
    union all
    select e.ra_id, e.title, e.venue_id, e.content_url, e.source, e.attending, 'room', 3
      from public.ra_events e
     where tr.show_date is not null and tr.show_venue_id is not null
       and e.event_date = tr.show_date and e.venue_id = tr.show_venue_id
       and (select count(*) from public.ra_events e2
             where e2.event_date = tr.show_date and e2.venue_id = tr.show_venue_id) = 1
  ) c
  order by c.pri, c.ra_id
  limit 1
) ev on true
left join public.venues v on v.id = coalesce(ev.venue_id, tr.show_venue_id);

-- Admin-only, like the tables under it. Never anon.
revoke all on public.v_radio_track_shows from public, anon;
grant select on public.v_radio_track_shows to authenticated;

commit;
