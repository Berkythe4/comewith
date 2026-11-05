-- 230: store the scan summary instead of computing it on every read.
--
-- 229's view derived top/median plays and the latest upload from `songs` at read
-- time. Correct (0 mismatches against the dashboard's JS over all 4,205 rows), but
-- the laterals ran over every row on every page: ~0.8s a page, 3.9s for the five
-- pages the Radio tab reads - so the 19x smaller download bought almost no time on
-- a fast line. The summary only changes when `songs` does, which is a scan.
--
-- So: five columns on sc_artist_cache, filled by a BEFORE trigger on any write that
-- touches `songs` (sc-enrich's upsert always sends it), and v_sc_artist_summary
-- reads them. Same view columns, same order, same types - the dashboard is unchanged.
-- The derivation rules are 229's, moved verbatim into the trigger; they must still
-- match scArtistData() in dashboard.html.

begin;

alter table public.sc_artist_cache
  add column if not exists top_title      text,
  add column if not exists top_url        text,
  add column if not exists top_plays      numeric not null default 0,
  add column if not exists med_plays      numeric not null default 0,
  add column if not exists latest_song_at timestamptz;

create or replace function public.sc_artist_cache_summarize()
returns trigger language plpgsql set search_path = public as $$
declare s jsonb := coalesce(new.songs, '[]'::jsonb);
begin
  if jsonb_typeof(s) <> 'array' then s := '[]'::jsonb; end if;
  new.top_title := null; new.top_url := null; new.top_plays := 0;
  -- Highest plays; ties to the first song, except ties at 0 go to the LAST
  -- (JS compares against `Number(best.playback_count) || -1`). See 229.
  select x->>'title', x->>'permalink_url', coalesce((x->>'playback_count')::numeric, 0)
    into new.top_title, new.top_url, new.top_plays
  from jsonb_array_elements(s) with ordinality as t(x, i)
  order by coalesce((x->>'playback_count')::numeric, 0) desc,
           case when coalesce((x->>'playback_count')::numeric, 0) = 0 then -i else i end asc
  limit 1;
  new.top_plays := coalesce(new.top_plays, 0);
  -- plays sorted ascending, element floor(n/2) (0-based); 0 when empty.
  select coalesce((
    select p.plays from (
      select coalesce((x->>'playback_count')::numeric, 0) as plays,
             row_number() over (order by coalesce((x->>'playback_count')::numeric, 0)) - 1 as rn,
             count(*) over () as n
      from jsonb_array_elements(s) x) p
    where p.rn = floor(p.n / 2)), 0)
    into new.med_plays;
  select max((x->>'created_at')::timestamptz) into new.latest_song_at
  from jsonb_array_elements(s) x where nullif(x->>'created_at', '') is not null;
  return new;
end $$;

revoke all on function public.sc_artist_cache_summarize() from public, anon;

drop trigger if exists sc_artist_cache_summarize on public.sc_artist_cache;
create trigger sc_artist_cache_summarize
  before insert or update of songs on public.sc_artist_cache
  for each row execute function public.sc_artist_cache_summarize();

-- Backfill: rewriting songs fires the trigger for every existing row.
update public.sc_artist_cache set songs = songs;

create or replace view public.v_sc_artist_summary
with (security_invoker = true) as
select
  c.soundcloud, c.sc_user_id, c.username, c.avatar_url, c.followers, c.city,
  c.is_producer, c.song_count, c.set_count, c.ok, c.scanned_at, c.sc_track_count,
  c.top_title, c.top_url, c.top_plays, c.med_plays, c.latest_song_at
from public.sc_artist_cache c;

revoke all on public.v_sc_artist_summary from public, anon;
grant select on public.v_sc_artist_summary to authenticated;

commit;
