-- =============================================================================
-- 109_traffic_views_and_external_radio.sql
-- 1. FIX: radio.episode_visits was counting Keith. All 35 logged visits belong
--    to berky@comewith.org (master_admin) — 18 on the live EP1 page and 17 on a
--    preview link — so the card read like an audience when it was one person
--    testing his own page. An audience KPI must mean EXTERNAL people, so both
--    radio metrics now exclude staff logins and preview URLs, and a new
--    radio.listeners card counts distinct outside listeners.
-- 2. Rollup views behind the Site traffic panel (top pages / links / referrers).
-- All views anon-revoked; nothing here grants anon anything.
-- =============================================================================
begin;

-- Who counts as an outside listener: a signed-in user who isn't staff.
-- deleted_at guard mirrors the 098 deactivation contract.
create or replace view public.v_radio_listeners as
select h.user_id,
       coalesce(p.role, 'customer')                                   as role,
       (p.role is null or p.role not in ('master_admin', 'sub_admin')) as is_external,
       sum(h.visits)                                                   as visits,
       count(*) filter (where h.station_slug not like 'preview:%')     as episodes_seen,
       max(h.last_seen_at)                                             as last_seen_at
  from public.listener_station_history h
  left join public.profiles p on p.id = h.user_id and p.deleted_at is null
 group by h.user_id, p.role;
revoke select on public.v_radio_listeners from anon;

-- Traffic rollups. 30-day windows to match the cards; the daily view stays
-- all-time so a longer chart is possible later without another migration.
create or replace view public.v_site_top_pages as
select path,
       count(*) filter (where kind = 'view')                       as views,
       count(distinct session_id) filter (where kind = 'view')     as visitors,
       count(*) filter (where kind = 'click')                      as clicks,
       max(occurred_at)                                            as last_seen
  from public.site_events
 where occurred_at >= now() - interval '30 days'
 group by path;
revoke select on public.v_site_top_pages from anon;

create or replace view public.v_site_top_links as
select link_url,
       min(link_label)                          as label,
       bool_or(outbound)                        as outbound,
       count(*)                                 as clicks,
       count(distinct session_id)               as clickers,
       max(occurred_at)                         as last_clicked
  from public.site_events
 where kind = 'click' and link_url is not null
   and occurred_at >= now() - interval '30 days'
 group by link_url;
revoke select on public.v_site_top_links from anon;

-- Where people came from. 'direct / none' is the honest bucket for a typed URL,
-- an app with no referrer, or a link out of a DM — not the same as unknown.
create or replace view public.v_site_referrers as
select coalesce(nullif(referrer, ''), 'direct / none')            as referrer,
       count(*) filter (where kind = 'view')                      as views,
       -- Must be filtered to views like the count beside it. Counting distinct
       -- sessions across ALL rows made a source read 1 view / 3 visitors,
       -- because click rows carry a session but no referrer.
       count(distinct session_id) filter (where kind = 'view')     as visitors,
       max(occurred_at)                                           as last_seen
  from public.site_events
 where occurred_at >= now() - interval '30 days'
 group by 1;
revoke select on public.v_site_referrers from anon;

-- ---------------------------------------------------------------------------
-- Recompute the KPI values with the radio metrics fixed.
-- ---------------------------------------------------------------------------
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
-- Outside listeners only, real episodes only (a preview link is Keith checking
-- his own work before it goes live — never an audience number).
rl as (
  select h.*
    from public.listener_station_history h
    left join public.profiles p on p.id = h.user_id
   where h.station_slug not like 'preview:%'
     and (p.role is null or p.role not in ('master_admin', 'sub_admin'))
)
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
  ('site.visitors_30d', (select count(distinct session_id)::numeric from se where kind = 'view')),
  ('site.views_30d',    (select count(*)::numeric from se where kind = 'view')),
  ('site.clicks_30d',   (select count(*)::numeric from se where kind = 'click')),
  ('guest.repeat_pct',     (select case when guests_with_attendance > 0 then round(100.0 * repeat_guests / guests_with_attendance, 1) end from gk))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('radio.listeners', 'audience', 'Radio listeners', 100, 'gte', '', current_date, true)
on conflict do nothing;

-- The two radio labels now promise "outside", because that's what they count.
update public.kpi_targets set label = 'Radio episode visits (outside)' where metric_key = 'radio.episode_visits';
update public.kpi_targets set label = 'Radio playlists saved (outside)' where metric_key = 'radio.playlists_saved';

commit;
