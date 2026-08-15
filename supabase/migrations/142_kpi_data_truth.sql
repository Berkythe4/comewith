-- =============================================================================
-- 142_kpi_data_truth.sql   (Strategy rebuild — Phase 1: make the data honest)
--
-- WHY: every live-COMPUTED KPI card (parties.*, di.*, audience.subscribers*,
-- guest.repeat_pct, radio.*, site.*) has always rendered "- no prior reading".
-- v_kpi_dashboard took current_value from coalesce(computed, snapshot) but
-- prior_value ONLY from v_metric_prior — the second-latest HAND-LOGGED
-- snapshot — and nobody hand-logs net P&L or subscriber counts. Introspection
-- on prod 2026-08-15 confirmed it: metric_snapshots holds readings for
-- youtube.* / instagram.* / tiktok.* and NOTHING else. So the metrics that
-- matter most are exactly the ones with no trend, and `as_of` for a computed
-- card was hardcoded to CURRENT_DATE, i.e. "updated today" forever.
--
-- This migration fixes the data layer only. It deliberately changes NOTHING on
-- the deployed dashboard:
--   * no new kpi_targets rows  -> no new cards appear on the live board
--   * no workstream re-categorisation -> radio.*/site.* stay in 'audience', so
--     the deployed renderer (which only knows content/audience/parties/
--     dance_infusion) cannot silently drop 9 cards.
-- Both of those land in Phase 2, together with the UI that organises them.
--
-- The one visible change: prior_value starts meaning "about 30 days ago"
-- instead of "the previous hand-logged reading". That is the fix; Phase 2 puts
-- the timeframe on the card face so the number says what it is.
--
-- GRANTS: 013_grants.sql sets `alter default privileges ... grant all on tables
-- to anon`, so EVERY new view here is auto-granted to anon on creation and must
-- be explicitly revoked. These are financial-adjacent (event P&L, money raised)
-- and follow decision E1 — anon must get 401.
-- =============================================================================
begin;

-- ---------------------------------------------------------------------------
-- 1. Per-event series for the event metrics.
--    Powers two things at once: the "last vs previous event" comparison (seq
--    1 vs 2) and the per-event bar charts Phase 2 draws. Completed events only,
--    same rule v_kpi_computed already uses -- an upcoming event has no P&L yet.
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_event_series as
with pt as (
  select k.event_id, k.name, k.event_date, k.net_pl, k.sell_through_pct
    from public.v_kpi_parties k
    join public.events e on e.id = k.event_id
   where e.status = 'completed'
), di as (
  select k.event_id, k.name, k.event_date, k.total_raised,
         k.cost_to_raise_per_dollar, k.total_attendance
    from public.v_kpi_dance_infusion k
    join public.events e on e.id = k.event_id
   where e.status = 'completed'
), s as (
  select 'parties.net_pl_last'::text       as metric_key, event_id, name, event_date, net_pl                     as value from pt
  union all
  select 'parties.sell_through_last'::text, event_id, name, event_date, sell_through_pct                          from pt
  union all
  select 'di.raised_last'::text,            event_id, name, event_date, total_raised                              from di
  union all
  select 'di.cost_to_raise_last'::text,     event_id, name, event_date, cost_to_raise_per_dollar                  from di
  union all
  select 'di.attendance_last'::text,        event_id, name, event_date, total_attendance::numeric                 from di
)
select s.metric_key, s.event_id, s.name, s.event_date, s.value,
       -- seq 1 = most recent completed event, 2 = the one before it
       row_number() over (partition by s.metric_key
                          order by s.event_date desc, s.event_id) as seq
  from s
 where s.value is not null;
revoke select on public.v_kpi_event_series from anon;

-- ---------------------------------------------------------------------------
-- 2. Recent content performance.
--    youtube.avg_views is LIFETIME views / all videos -- its own tooltip admits
--    it cannot read recent performance. This is the last 5 uploads against the
--    5 before them, which is the number that tells you whether what you are
--    making NOW is landing.
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_content_recent as
with ranked as (
  select views,
         row_number() over (order by published_at desc nulls last) as rn
    from public.youtube_videos
)
select (select round(avg(views), 0) from ranked where rn <= 5)              as avg_views_recent,
       (select round(avg(views), 0) from ranked where rn between 6 and 10)  as avg_views_prior;
revoke select on public.v_kpi_content_recent from anon;

-- ---------------------------------------------------------------------------
-- 3. v_kpi_computed -- unchanged body, plus the new keys.
--    Reproduced from the LIVE prod definition (pg_get_viewdef, 2026-08-15), not
--    from the migration files: 051/107/108/110 each layered onto it, so the
--    files alone are not the truth.
--    The new keys are computed but NOT carded -- a value here is inert until a
--    kpi_targets row exists, which is Phase 2's job.
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_computed as
with di as (
  select k.event_id, k.name, k.event_date, k.total_attendance, k.net_pl,
         k.total_raised, k.cost_to_raise_per_dollar
    from public.v_kpi_dance_infusion k
    join public.events e on e.id = k.event_id
   where e.status = 'completed'
), pt as (
  select k.event_id, k.name, k.event_date, k.capacity, k.tickets_sold,
         k.sell_through_pct, k.net_pl
    from public.v_kpi_parties k
    join public.events e on e.id = k.event_id
   where e.status = 'completed'
), gk as (
  select * from public.v_guest_kpis limit 1
), se as (
  select * from public.site_events where occurred_at >= now() - interval '30 days'
), rl as (
  select h.*
    from public.listener_station_history h
    left join public.profiles p on p.id = h.user_id
   where h.station_slug not like 'preview:%'
     and (p.role is null or p.role <> all (array['master_admin', 'sub_admin']))
), rm as (
  select * from public.v_radio_media
), rmy as (
  select * from rm where rm.yt_views is not null
), es as (
  select * from public.v_kpi_event_series where seq = 1
), cr as (
  select * from public.v_kpi_content_recent
)
select metric_key, value from (values
  -- ---- existing keys, byte-for-byte the same arithmetic as before ----------
  ('di.raised_per_event',   (select round(avg(total_raised), 2) from di)),
  ('di.cost_to_raise',      (select round(avg(cost_to_raise_per_dollar), 2) from di)),
  ('di.attendance',         (select round(avg(total_attendance), 0) from di)),
  ('di.to_ms_total',        (select sum(net_pl) from di)),
  ('parties.net_pl',        (select round(avg(net_pl), 2) from pt)),
  ('parties.sell_through',  (select round(avg(sell_through_pct), 1) from pt)),
  ('parties.net_pl_total',  (select sum(net_pl) from pt)),
  ('audience.subscribers',  (select count(*)::numeric from public.subscribers where status = 'subscribed')),
  ('audience.subscribers_come_with', (select count(distinct s.id)::numeric
      from public.subscribers s join public.subscriber_segments g on g.subscriber_id = s.id
     where s.status = 'subscribed' and g.segment = 'come_with')),
  ('audience.subscribers_dance_infusion', (select count(distinct s.id)::numeric
      from public.subscribers s join public.subscriber_segments g on g.subscriber_id = s.id
     where s.status = 'subscribed' and g.segment = 'dance_infusion')),
  ('radio.listeners',       (select count(distinct user_id)::numeric from rl)),
  ('radio.episode_visits',  (select coalesce(sum(visits), 0)::numeric from rl)),
  ('radio.playlists_saved', (select count(*)::numeric
      from public.listener_playlists l left join public.profiles p on p.id = l.user_id
     where p.role is null or p.role <> all (array['master_admin', 'sub_admin']))),
  ('radio.yt_avg_views',    (select round(avg(yt_views), 0) from rmy)),
  ('radio.yt_views_total',  (select sum(yt_views) from rmy)),
  ('radio.yt_engagement',   (select case when sum(yt_views) > 0
      then round(100.0 * (sum(yt_likes) + sum(yt_comments)) / sum(yt_views), 2) end from rmy)),
  ('radio.sc_plays',        (select sum(mix_sc_plays)::numeric from rm)),
  ('site.visitors_30d',     (select count(distinct session_id)::numeric from se where kind = 'view')),
  ('site.views_30d',        (select count(*)::numeric from se where kind = 'view')),
  ('site.clicks_30d',       (select count(*)::numeric from se where kind = 'click')),
  ('guest.repeat_pct',      (select case when guests_with_attendance > 0
      then round(100.0 * repeat_guests::numeric / guests_with_attendance::numeric, 1) end from gk)),
  -- ---- NEW: last-event values (the health headline reads these) -----------
  ('parties.net_pl_last',       (select value from es where metric_key = 'parties.net_pl_last')),
  ('parties.sell_through_last', (select value from es where metric_key = 'parties.sell_through_last')),
  ('di.raised_last',            (select value from es where metric_key = 'di.raised_last')),
  ('di.cost_to_raise_last',     (select value from es where metric_key = 'di.cost_to_raise_last')),
  ('di.attendance_last',        (select value from es where metric_key = 'di.attendance_last')),
  -- ---- NEW: recent content, not the lifetime average ----------------------
  ('content.avg_views_recent',  (select avg_views_recent from cr))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

-- ---------------------------------------------------------------------------
-- 4. v_metric_latest gains `source`, so a card can say where its number came
--    from (computed / youtube_api / manual) instead of hiding it in a tooltip.
--    Appended at the END -- create-or-replace cannot reorder or rename.
-- ---------------------------------------------------------------------------
create or replace view public.v_metric_latest as
select distinct on (metric_key, series_id)
  metric_key, series_id, value, captured_on, source
from public.metric_snapshots
order by metric_key, series_id, captured_on desc;
revoke select on public.v_metric_latest from anon;

-- ---------------------------------------------------------------------------
-- 5. What "prior" means, per metric, in ONE place.
--    (a) event metrics  -> the PREVIOUS completed event. "vs the last party"
--        is the comparison Keith actually makes; a calendar window would
--        straddle a gap where nothing happened.
--    (b) recent content -> the 5 uploads before the latest 5.
--    (c) everything else -> the reading nearest 30 days ago. Falls back to the
--        EARLIEST reading on record when history is shorter than 30 days --
--        falling back to the latest instead would compare a number to itself
--        and render a permanent "no change".
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_prior as
select metric_key, value as prior_value, event_date as prior_as_of,
       'previous event'::text as prior_basis
  from public.v_kpi_event_series
 where seq = 2
union all
select 'content.avg_views_recent', avg_views_prior, null::date, 'previous 5 uploads'
  from public.v_kpi_content_recent
 where avg_views_prior is not null
union all
-- Parenthesised: a bare ORDER BY on a UNION branch binds to the whole union,
-- and DISTINCT ON needs its own ordering.
(select distinct on (metric_key)
       metric_key, value, captured_on,
       case when captured_on <= current_date - 30 then 'about 30 days ago'
            else 'earliest reading on record' end
  from public.metric_snapshots
 where series_id is null
   and metric_key not in (select metric_key from public.v_kpi_event_series)
   and metric_key <> 'content.avg_views_recent'
 order by metric_key,
          -- prefer a reading at least 30 days old...
          (captured_on <= current_date - 30) desc,
          -- ...the most recent of those; otherwise the earliest we have
          case when captured_on <= current_date - 30 then captured_on end desc nulls last,
          captured_on asc);
revoke select on public.v_kpi_prior from anon;

-- ---------------------------------------------------------------------------
-- 6. When did this number last actually MOVE?
--    With a nightly snapshot, "last captured" is always today and means
--    nothing. This walks the history and returns the start of the current run
--    of identical values, so a card can say "unchanged since Jul 2".
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_changed as
with h as (
  -- lag() must resolve in its own scope: a window call cannot be nested
  -- inside another window call's argument.
  select metric_key, value, captured_on,
         lag(value) over (partition by metric_key order by captured_on) as prev
    from public.metric_snapshots
   where series_id is null
), marked as (
  select metric_key, value, captured_on,
         sum(case when value is distinct from prev then 1 else 0 end)
           over (partition by metric_key order by captured_on) as grp
    from h
)
select distinct on (metric_key)
       metric_key,
       min(captured_on) over (partition by metric_key, grp) as changed_on
  from marked
 order by metric_key, captured_on desc;
revoke select on public.v_kpi_changed from anon;

-- ---------------------------------------------------------------------------
-- 7. The board. First 9 columns are unchanged in name and order (create-or-
--    replace forbids otherwise); the new ones are appended.
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_dashboard as
select t.metric_key, t.workstream, t.label,
       coalesce(c.value, l.value)        as current_value,
       pr.prior_value                    as prior_value,
       t.target_value, t.comparison, t.unit,
       -- honest: the date of the reading behind the number, not CURRENT_DATE
       coalesce(l.captured_on, current_date) as as_of,
       -- appended --------------------------------------------------------
       pr.prior_as_of,
       pr.prior_basis,
       case when c.value is not null      then 'computed'
            when l.source = 'youtube_api' then 'api'
            when l.value is not null      then 'manual'
            end                           as source_kind,
       ch.changed_on
  from public.v_kpi_targets_current t
  left join public.v_kpi_computed  c  on c.metric_key  = t.metric_key
  left join public.v_metric_latest l  on l.metric_key  = t.metric_key and l.series_id is null
  left join public.v_kpi_prior     pr on pr.metric_key = t.metric_key
  left join public.v_kpi_changed   ch on ch.metric_key = t.metric_key;
revoke select on public.v_kpi_dashboard from anon;

-- ---------------------------------------------------------------------------
-- 8. The nightly snapshot -- the thing that makes every chart above possible.
--    Writes v_kpi_computed into metric_snapshots with source='computed', so
--    computed metrics build the same history the YouTube job already builds.
--    Idempotent: re-running on the same day updates that day's reading.
-- ---------------------------------------------------------------------------
create or replace function public.snapshot_kpis()
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare n integer;
begin
  insert into public.metric_snapshots (metric_key, value, captured_on, series_id, source, notes)
  select c.metric_key, c.value, current_date, null, 'computed', 'nightly KPI snapshot'
    from public.v_kpi_computed c
   where c.value is not null
      on conflict (metric_key, captured_on,
                   coalesce(series_id, '00000000-0000-0000-0000-000000000000'::uuid))
      do update set value = excluded.value, source = 'computed';
  get diagnostics n = row_count;
  return n;
end
$$;
-- 013's default privileges grant new functions to anon -- take it back.
revoke all on function public.snapshot_kpis() from anon;

-- 06:30 UTC: after pull-youtube-stats (06:00), so the day's YouTube numbers are
-- already in when the snapshot runs. Re-running returns the same job id.
select cron.schedule('snapshot-kpis', '30 6 * * *', $$select public.snapshot_kpis()$$);

-- Seed day one, so history starts now rather than tomorrow morning.
select public.snapshot_kpis();

-- ---------------------------------------------------------------------------
-- 9. Per-user dashboard prefs.
--    dashboard_prefs is a SINGLETON row, so one person hiding a card hides it
--    for everyone -- unworkable once categories collapse per person. The old
--    table stays untouched: the deployed dashboard still reads it, and Phase 2
--    moves the UI over.
--    Storing EXPANDED (not collapsed) means the page defaults to the summary
--    view, which is the whole point of the redesign. Row-local predicate in the
--    SELECT policy per CLAUDE.md, so insert..returning works for the owner.
-- ---------------------------------------------------------------------------
create table if not exists public.user_dashboard_prefs (
  user_id              uuid primary key references auth.users(id) on delete cascade,
  hidden_metric_keys   text[] not null default '{}',
  expanded_categories  text[] not null default '{}',
  updated_at           timestamptz not null default now()
);
alter table public.user_dashboard_prefs enable row level security;
drop policy if exists "Users manage their own dashboard prefs" on public.user_dashboard_prefs;
create policy "Users manage their own dashboard prefs"
  on public.user_dashboard_prefs for all
  using (user_id = auth.uid())
  with check (user_id = auth.uid());
revoke all on public.user_dashboard_prefs from anon;

-- ---------------------------------------------------------------------------
-- 10. Duplicate active targets.
--     prod has youtube.avg_views x4 (3999/4000/499/500) and di.cost_to_raise x2
--     (0.25/0.50). v_kpi_targets_current already dedups with DISTINCT ON, so
--     only one was ever displayed -- but which one is invisible to the reader,
--     and cost-to-raise is now the Dance Infusion health headline, so an
--     ambiguous target is not acceptable.
--     This deactivates every active row EXCEPT the one the view already picks,
--     so the displayed targets do not change by a single digit.
-- ---------------------------------------------------------------------------
with winners as (
  select distinct on (metric_key) id
    from public.kpi_targets
   where active
   order by metric_key, effective_date desc, updated_at desc, id desc
)
update public.kpi_targets
   set active = false, updated_at = now()
 where active
   and id not in (select id from winners);

commit;

-- DOWN:
--   select cron.unschedule('snapshot-kpis');
--   drop function public.snapshot_kpis();
--   drop table public.user_dashboard_prefs;
--   drop view public.v_kpi_changed, public.v_kpi_prior,
--             public.v_kpi_content_recent, public.v_kpi_event_series;
--   recreate v_kpi_dashboard / v_kpi_computed / v_metric_latest from the
--   pre-142 definitions captured in the session notes;
--   (the kpi_targets deactivations are data -- re-activate by hand if needed).
