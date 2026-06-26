-- 063_recap_videos.sql
-- Multiple recap videos per event, each with its own label (the text shown on
-- the public "watch the recap" link). Replaces the single events.youtube_url as
-- the source of truth; youtube_url is kept in sync (first video) for the card
-- thumbnail + back-compat.

alter table public.events
  add column if not exists recap_videos jsonb not null default '[]'::jsonb;

-- Migrate any existing single youtube_url into the array.
update public.events
set recap_videos = jsonb_build_array(
      jsonb_build_object('url', youtube_url, 'label', 'Watch the recap'))
where youtube_url is not null and youtube_url <> ''
  and (recap_videos is null or recap_videos = '[]'::jsonb);

-- Expose recap_videos to the public recap view (anon grant is preserved by
-- create-or-replace).
create or replace view public.v_public_recap as
  select e.id,
         e.name,
         e.event_date,
         v.name as venue_name,
         e.series,
         e.type,
         e.hero_image_path,
         e.youtube_url,
         e.recap_blurb,
         e.recap_videos
  from public.events e
    left join public.venues v on v.id = e.venue_id
  where e.is_featured = true and e.deleted_at is null
  order by e.event_date desc;

notify pgrst, 'reload schema';
