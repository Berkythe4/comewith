-- =============================================================================
-- 220_radio_post_classified.sql
-- The scheduled go-live (radio_publish_station) drops a "posted" release card on
-- the social calendar. Since 217 that card landed with the column defaults -
-- phase general, account come_with, no format - so every release needed hand
-- re-classifying (219 did the backlog). Now it is born phase radio / come_with /
-- reel. Taken from the LIVE definition (pg_get_functiondef, 2026-10-02); only the
-- social_posts insert changed. sc-connect's manual Go live gets the same three
-- fields in the same commit (supabase/functions/sc-connect/radio_post.ts).
-- No table or column changes.
-- =============================================================================
begin;

CREATE OR REPLACE FUNCTION public.radio_publish_station(p_id uuid)
 RETURNS text
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
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
    -- 220: the release card is classified at birth - phase radio, the Come With
    -- account, and a reel (a release is the episode video; recaps are carousels
    -- and are planned by hand / by the connector, never auto-created here).
    insert into social_posts (title, caption, channels, series, content_pillar, stage,
                              scheduled_for, posted_at, link_url, phase, account, format)
    values (btrim('📻 Come With Radio SHOW ' || coalesce(v_pl.station_no::text,'') || ' — ' || coalesce(v_pl.name,'')),
            left(coalesce(v_pl.desc_sc,''), 1000), array['other'], 'Come With Radio', 'radio episode',
            'posted', v_now, v_now, 'https://comewith.org/radio.html?s=' || v_slug,
            'radio', 'come_with', 'reel');
  exception when others then null;
  end;

  return v_slug;
end;
$function$;

commit;
