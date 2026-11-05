-- 229: v_sc_artist_summary - the SoundCloud scan cache WITHOUT the song lists.
--
-- Why: opening the Radio tab downloaded sc_artist_cache whole - 3,911 rows, 32 MB
-- of JSON (6.7 MB gzipped, ~5s on a fast line), four sequential pages - every time.
-- Nearly all of it is `songs`, and the only things the list views need from it are
-- three numbers for Buzz (top-track plays, typical-track plays, released in the last
-- 6 months) plus the top track's title/link. Measured without songs: 0.9 MB, 1.2s.
--
-- The dashboard now reads THIS on the Radio tab, Market and Scene Map, and loads full
-- song lists only when the Build view is opened (for the artists in its window).
--
-- The derived fields MUST match scArtistData() in dashboard.html, which still
-- computes them from `songs` when they are loaded, and dj-station (which reads songs
-- server-side). The rules, matched exactly:
--   top      = song with the highest playback_count (JS reduce; null counts as 0);
--              ties go to the first song, except ties at 0 (see the ORDER BY).
--   med      = plays sorted ascending, element floor(n/2) (0-based) - NOT a true
--              median and NOT percentile_disc; 0 when there are no songs.
--   latest   = the newest created_at; JS tests "any created_at >= now - 182 days",
--              which is the same as "the newest one is".
-- security_invoker so sc_artist_cache's is_admin() RLS applies to the view too.

begin;

create or replace view public.v_sc_artist_summary
with (security_invoker = true) as
select
  c.soundcloud, c.sc_user_id, c.username, c.avatar_url, c.followers, c.city,
  c.is_producer, c.song_count, c.set_count, c.ok, c.scanned_at, c.sc_track_count,
  top.title                                     as top_title,
  top.permalink_url                             as top_url,
  coalesce(top.plays, 0)                        as top_plays,
  coalesce(med.plays, 0)                        as med_plays,
  lat.latest_song_at
from public.sc_artist_cache c
left join lateral (
  select s->>'title' as title, s->>'permalink_url' as permalink_url,
         coalesce((s->>'playback_count')::numeric, 0) as plays
  from jsonb_array_elements(coalesce(c.songs, '[]'::jsonb)) with ordinality as t(s, i)
  -- JS compares against `Number(best.playback_count) || -1`, so a best with 0 plays
  -- reads as -1 and the NEXT zero replaces it: ties at 0 go to the LAST song, ties
  -- above 0 to the FIRST. Only the displayed top title depends on this.
  order by coalesce((s->>'playback_count')::numeric, 0) desc,
           case when coalesce((s->>'playback_count')::numeric, 0) = 0 then -i else i end asc
  limit 1
) top on true
left join lateral (
  select p.plays
  from (
    select coalesce((s->>'playback_count')::numeric, 0) as plays,
           row_number() over (order by coalesce((s->>'playback_count')::numeric, 0)) - 1 as rn,
           count(*) over () as n
    from jsonb_array_elements(coalesce(c.songs, '[]'::jsonb)) s
  ) p
  where p.rn = floor(p.n / 2)
) med on true
left join lateral (
  select max((s->>'created_at')::timestamptz) as latest_song_at
  from jsonb_array_elements(coalesce(c.songs, '[]'::jsonb)) s
  where nullif(s->>'created_at', '') is not null
) lat on true;

-- Admin-only, like the table. Never anon.
revoke all on public.v_sc_artist_summary from public, anon;
grant select on public.v_sc_artist_summary to authenticated;

commit;
