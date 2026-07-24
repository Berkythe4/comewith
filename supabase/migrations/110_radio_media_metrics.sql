-- =============================================================================
-- 110_radio_media_metrics.sql
-- Radio is the audience-building bet, so the episode MEDIA has to be measurable
-- next to the listener counts: how the mix performs on YouTube and SoundCloud,
-- not just how many people signed in.
--   * YouTube is already in hand — pull-youtube-stats fills youtube_videos for
--     the whole channel, so an episode's numbers are a join on the video id
--     pulled out of sc_playlists.mix_youtube_url. Nothing new to fetch.
--   * SoundCloud play counts were stored NOWHERE. Columns added here, filled by
--     the new sc-connect "mix_stats" action.
-- =============================================================================
begin;

alter table public.sc_playlists add column if not exists mix_sc_plays    int;
alter table public.sc_playlists add column if not exists mix_sc_likes    int;
alter table public.sc_playlists add column if not exists mix_sc_reposts  int;
alter table public.sc_playlists add column if not exists mix_sc_comments int;
alter table public.sc_playlists add column if not exists mix_stats_at    timestamptz;

-- One row per published episode with both platforms side by side.
-- The video id is extracted from whatever URL shape got pasted: watch?v=…,
-- youtu.be/…, or /embed/… — a stored URL is a human artefact, not a clean id.
create or replace view public.v_radio_media as
select p.id,
       p.station_no,
       p.name,
       p.published_at,
       p.mix_youtube_url,
       coalesce(
         substring(p.mix_youtube_url from '[?&]v=([A-Za-z0-9_-]{6,})'),
         substring(p.mix_youtube_url from 'youtu\.be/([A-Za-z0-9_-]{6,})'),
         substring(p.mix_youtube_url from '/embed/([A-Za-z0-9_-]{6,})')
       )                                            as yt_video_id,
       v.title                                      as yt_title,
       v.views                                      as yt_views,
       v.likes                                      as yt_likes,
       v.comments                                   as yt_comments,
       p.mix_sc_plays, p.mix_sc_likes, p.mix_sc_reposts, p.mix_sc_comments,
       p.mix_stats_at
  from public.sc_playlists p
  left join public.youtube_videos v
         on v.video_id = coalesce(
              substring(p.mix_youtube_url from '[?&]v=([A-Za-z0-9_-]{6,})'),
              substring(p.mix_youtube_url from 'youtu\.be/([A-Za-z0-9_-]{6,})'),
              substring(p.mix_youtube_url from '/embed/([A-Za-z0-9_-]{6,})'))
 where p.status = 'live';
revoke select on public.v_radio_media from anon;

create or replace view public.v_kpi_computed as
with di as (
  select k.* from public.v_kpi_dance_infusion k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
pt as (
  select k.* from public.v_kpi_parties k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
gk as (select * from public.v_guest_kpis limit 1),
se as (select * from public.site_events where occurred_at >= now() - interval '30 days'),
rl as (
  select h.*
    from public.listener_station_history h
    left join public.profiles p on p.id = h.user_id
   where h.station_slug not like 'preview:%'
     and (p.role is null or p.role not in ('master_admin', 'sub_admin'))
),
-- Episode media. yt rows only count episodes whose video was actually found in
-- the channel pull — an episode with no video must not drag an average to zero.
rm  as (select * from public.v_radio_media),
rmy as (select * from rm where yt_views is not null)
select metric_key, value from (values
  ('di.raised_per_event',  (select round(avg(total_raised), 2) from di)),
  ('di.cost_to_raise',     (select round(avg(cost_to_raise_per_dollar), 2) from di)),
  ('di.attendance',        (select round(avg(total_attendance), 0) from di)),
  ('di.to_ms_total',       (select sum(net_pl) from di)),
  ('parties.net_pl',       (select round(avg(net_pl), 2) from pt)),
  ('parties.sell_through', (select round(avg(sell_through_pct), 1) from pt)),
  ('parties.net_pl_total', (select sum(net_pl) from pt)),
  ('audience.subscribers', (select count(*)::numeric from public.subscribers where status = 'subscribed')),
  ('audience.subscribers_come_with', (
     select count(distinct s.id)::numeric from public.subscribers s
       join public.subscriber_segments g on g.subscriber_id = s.id
      where s.status = 'subscribed' and g.segment = 'come_with')),
  ('audience.subscribers_dance_infusion', (
     select count(distinct s.id)::numeric from public.subscribers s
       join public.subscriber_segments g on g.subscriber_id = s.id
      where s.status = 'subscribed' and g.segment = 'dance_infusion')),
  ('radio.listeners',      (select count(distinct user_id)::numeric from rl)),
  ('radio.episode_visits', (select coalesce(sum(visits), 0)::numeric from rl)),
  ('radio.playlists_saved', (
     select count(*)::numeric from public.listener_playlists l
       left join public.profiles p on p.id = l.user_id
      where p.role is null or p.role not in ('master_admin', 'sub_admin'))),
  ('radio.yt_avg_views',   (select round(avg(yt_views), 0) from rmy)),
  ('radio.yt_views_total', (select sum(yt_views)::numeric from rmy)),
  -- Engagement pooled across episodes (total interactions ÷ total views), not
  -- an average of per-episode rates: one 3-view episode shouldn't swing it.
  ('radio.yt_engagement',  (select case when sum(yt_views) > 0
                                   then round(100.0 * (sum(yt_likes) + sum(yt_comments)) / sum(yt_views), 2) end from rmy)),
  ('radio.sc_plays',       (select sum(mix_sc_plays)::numeric from rm)),
  ('site.visitors_30d', (select count(distinct session_id)::numeric from se where kind = 'view')),
  ('site.views_30d',    (select count(*)::numeric from se where kind = 'view')),
  ('site.clicks_30d',   (select count(*)::numeric from se where kind = 'click')),
  ('guest.repeat_pct',     (select case when guests_with_attendance > 0 then round(100.0 * repeat_guests / guests_with_attendance, 1) end from gk))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

-- Workstream stays 'audience' so these land on the radio row; the dashboard
-- paints them with the Content colour to mark them as media numbers.
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('radio.yt_avg_views',   'audience', 'Episode avg views (YT)',  500, 'gte', '',  current_date, true),
  ('radio.yt_views_total', 'audience', 'Episode views (YT)',     2000, 'gte', '',  current_date, true),
  ('radio.yt_engagement',  'audience', 'Episode engagement (YT)',   5, 'gte', '%', current_date, true),
  ('radio.sc_plays',       'audience', 'Mix plays (SoundCloud)',  1000, 'gte', '', current_date, true)
on conflict do nothing;

commit;
