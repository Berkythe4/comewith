-- =============================================================================
-- 106_scheduled_go_live.sql
-- Schedule an episode to publish itself at a set time, and open next week's
-- station the moment you schedule (so you can start building it while this one
-- waits to drop). DB-ONLY publish: mirrors the SQL side of sc-connect `finalize`
-- (publish page, log played, open next carry-over station, social card) but NOT
-- the SoundCloud API push. The manual 🚀 Go live button is unchanged.
--
--   scheduled_go_live timestamptz    "publish at this time"
--   radio_open_next_station()        create next week's carry-over station (or
--                                    return the existing working one) — shared
--   radio_schedule_go_live(id,when,slug)  admin RPC: lock + schedule + open next
--   radio_publish_station(id)        the SQL publish for one station
--   radio_publish_due()              publish everything whose time has come
--   pg_cron every 5 min → radio_publish_due()
--
-- One-working-station invariant is respected: scheduling flips the scheduled
-- station building→testing (still editable), so the freshly-opened next station
-- is the only 'building' one. Flip between any station via the dashboard switcher.
-- =============================================================================
begin;

alter table public.sc_playlists add column if not exists scheduled_go_live timestamptz;

-- Open next week's station (idempotent): if a working ('building') station
-- already exists, return it; otherwise create the next number and auto-carry
-- every passed-not-played-not-carried song.
create or replace function public.radio_open_next_station()
returns uuid
language plpgsql security definer set search_path = public
as $$
declare v_id uuid; v_no int; v_pos int := 0; v_now timestamptz := now(); c record;
begin
  select id into v_id from sc_playlists where status = 'building' order by station_no desc limit 1;
  if found then return v_id; end if;
  select coalesce(max(station_no), 0) + 1 into v_no from sc_playlists;
  insert into sc_playlists (name, station_no) values ('Weekly station', v_no) returning id into v_id;
  for c in
    select sc_track_id, title, artist_name, permalink_url, artwork_url, duration_ms, passed_station_no
      from sc_song_log
     where passed_at is not null and played_at is null and carried_to is null
     order by passed_at
  loop
    v_pos := v_pos + 10;
    begin
      insert into sc_playlist_tracks (playlist_id, sc_track_id, title, artist_name, permalink_url,
                                      artwork_url, duration_ms, sort, carried_from)
      values (v_id, c.sc_track_id, c.title, c.artist_name, c.permalink_url,
              c.artwork_url, c.duration_ms, v_pos, c.passed_station_no);
      update sc_song_log set carried_to = v_id, updated_at = v_now where sc_track_id = c.sc_track_id;
    exception when unique_violation then null;
    end;
  end loop;
  return v_id;
end;
$$;

-- Admin RPC: lock this station in as scheduled, and open next week's now.
create or replace function public.radio_schedule_go_live(p_id uuid, p_when timestamptz, p_slug text default null)
returns jsonb
language plpgsql security definer set search_path = public
as $$
declare v_pl record; v_slug text; v_next uuid;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;
  select * into v_pl from sc_playlists where id = p_id;
  if not found then raise exception 'station not found'; end if;
  if not exists (select 1 from sc_playlist_tracks where playlist_id = p_id) then
    raise exception 'this station has no tracks';
  end if;

  v_slug := coalesce(nullif(btrim(p_slug), ''), nullif(btrim(v_pl.slug), ''),
    left(regexp_replace(regexp_replace(lower(coalesce(v_pl.name,'station')), '[^a-z0-9]+','-','g'),
                        '(^-+|-+$)','','g'), 50) || '-ep' || coalesce(v_pl.station_no,0));
  if exists (select 1 from sc_playlists where slug = v_slug and id <> p_id) then
    v_slug := v_slug || '-' || substr(md5(p_id::text),1,4);
  end if;

  update sc_playlists
     set slug = v_slug, scheduled_go_live = p_when,
         status = case when status = 'building' then 'testing' else status end,
         updated_at = now()
   where id = p_id;

  v_next := public.radio_open_next_station();     -- start next week now
  return jsonb_build_object('slug', v_slug, 'next_station', v_next, 'scheduled_go_live', p_when);
end;
$$;

-- Publish ONE station (SQL half of finalize). Returns the slug, or null if it
-- couldn't (no slug/tracks, or already live).
create or replace function public.radio_publish_station(p_id uuid)
returns text
language plpgsql security definer set search_path = public
as $$
declare v_pl record; v_slug text; v_now timestamptz := now();
begin
  select * into v_pl from sc_playlists where id = p_id;
  if not found then return null; end if;
  if v_pl.status = 'live' or v_pl.published then return v_pl.slug; end if;
  if not exists (select 1 from sc_playlist_tracks where playlist_id = p_id) then return null; end if;

  v_slug := nullif(btrim(v_pl.slug), '');
  if v_slug is null then
    v_slug := left(regexp_replace(regexp_replace(lower(coalesce(v_pl.name,'station')), '[^a-z0-9]+','-','g'),
                                  '(^-+|-+$)','','g'), 50) || '-ep' || coalesce(v_pl.station_no,0);
  end if;
  if exists (select 1 from sc_playlists where slug = v_slug and id <> p_id) then
    v_slug := v_slug || '-' || substr(md5(p_id::text),1,4);
  end if;

  update sc_playlists
     set slug = v_slug, published = true, status = 'live',
         published_at = coalesce(published_at, v_now), scheduled_go_live = null, updated_at = v_now
   where id = p_id;

  insert into sc_song_log (sc_track_id, title, artist_name, permalink_url, artwork_url,
                           duration_ms, played_playlist_id, played_station_no, played_at, updated_at)
  select t.sc_track_id, t.title, t.artist_name, t.permalink_url, t.artwork_url,
         t.duration_ms, p_id, v_pl.station_no, v_now, v_now
    from sc_playlist_tracks t where t.playlist_id = p_id
  on conflict (sc_track_id) do update
     set played_playlist_id = excluded.played_playlist_id, played_station_no = excluded.played_station_no,
         played_at = excluded.played_at, updated_at = excluded.updated_at;

  perform public.radio_open_next_station();   -- no-op if scheduling already opened it

  begin
    insert into social_posts (title, caption, channels, series, content_pillar, stage,
                              scheduled_for, posted_at, link_url)
    values (btrim('📻 Come With Radio EP ' || coalesce(v_pl.station_no::text,'') || ' — ' || coalesce(v_pl.name,'')),
            left(coalesce(v_pl.desc_sc,''), 1000), array['other'], 'Come With Radio', 'radio episode',
            'posted', v_now, v_now, 'https://comewith.org/radio.html?s=' || v_slug);
  exception when others then null;
  end;

  return v_slug;
end;
$$;

create or replace function public.radio_publish_due()
returns int
language plpgsql security definer set search_path = public
as $$
declare r record; n int := 0;
begin
  for r in
    select id from sc_playlists
     where scheduled_go_live is not null and scheduled_go_live <= now()
       and status <> 'live' and published is not true
     order by scheduled_go_live
  loop
    begin
      if public.radio_publish_station(r.id) is not null then n := n + 1; end if;
    exception when others then raise warning 'radio_publish_due % failed: %', r.id, sqlerrm;
    end;
  end loop;
  return n;
end;
$$;

-- Functions default to EXECUTE for PUBLIC, so revoke from PUBLIC (not just anon)
-- or they stay callable via the REST RPC endpoint. The cron functions are
-- cron-only (owner runs them); only the admin-gated schedule RPC is exposed.
revoke execute on function public.radio_open_next_station() from public, anon, authenticated;
revoke execute on function public.radio_publish_station(uuid) from public, anon, authenticated;
revoke execute on function public.radio_publish_due() from public, anon, authenticated;
revoke execute on function public.radio_schedule_go_live(uuid, timestamptz, text) from public, anon;
grant execute on function public.radio_schedule_go_live(uuid, timestamptz, text) to authenticated;

do $$ begin perform cron.unschedule('radio-publish-due'); exception when others then null; end $$;
select cron.schedule('radio-publish-due', '*/5 * * * *', $$select public.radio_publish_due()$$);

commit;
-- POST: sc_playlists.scheduled_go_live; radio_open_next_station /
-- radio_schedule_go_live (admin RPC) / radio_publish_station / radio_publish_due
-- (security definer, anon-revoked); pg_cron 'radio-publish-due' every 5 min.
