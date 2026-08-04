-- 137: the global counter is a SHOW number, not an EP number.
--
-- Two numbers had been sharing one word. `sc_playlists.station_no` counts every
-- broadcast we have ever put out; `edition_seq` counts episodes WITHIN a special
-- edition. While the NYC weekly was the only series those were the same number,
-- so both were rendered "EP n". The Elements edition broke that: its Ep1 is the
-- 4th show we have made, so the Control Center read "EP 4 · Come With Elements
-- Radio — Ep1" and the social post would have gone out titled "EP 4" for an
-- episode the world knows as Elements Ep1.
--
-- So the global counter is now called SHOW everywhere it is displayed, and
-- "episode" is reserved for a series' own numbering. This migration carries that
-- wording into the two places it is baked into the DATABASE — the rest is
-- front-end (dashboard/radio/dj/index) and sc-connect.
--
-- Text-only. No schema, no data, no policy change. Both functions are replaced
-- with their current prod definition plus the reworded string; introspected
-- against prod before writing (only these two contain an 'EP ' literal).

-- 1) The scheduled-release path's auto social post. The edge function
--    (sc-connect finalize) already says SHOW; this is the SQL half that
--    radio-publish-due / the cron backstop go through, and the two must agree.
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
    values (btrim('📻 Come With Radio SHOW ' || coalesce(v_pl.station_no::text,'') || ' — ' || coalesce(v_pl.name,'')),
            left(coalesce(v_pl.desc_sc,''), 1000), array['other'], 'Come With Radio', 'radio episode',
            'posted', v_now, v_now, 'https://comewith.org/radio.html?s=' || v_slug);
  exception when others then null;
  end;

  return v_slug;
end;
$$;

-- 2) The closed-episode guard's error text (135). Rarely seen — the dashboard
--    catches this case first — but when it does surface it should name the show
--    the same way every other surface does.
create or replace function public.sc_tracks_block_closed()
returns trigger
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  st text;
  no int;
begin
  select status, station_no into st, no
    from public.sc_playlists where id = new.playlist_id;

  if st in ('live', 'archived') then
    raise exception
      'SHOW % is % — reopen the episode before adding songs (set its status back to testing).',
      coalesce(no::text, '?'), st
      using errcode = 'check_violation';
  end if;

  return new;
end;
$$;

comment on function public.sc_tracks_block_closed() is
  'Blocks INSERTs (and re-parenting UPDATEs) of tracks onto live/archived episodes. See migration 135.';
