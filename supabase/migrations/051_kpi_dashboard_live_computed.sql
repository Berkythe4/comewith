-- =============================================================================
-- 051_kpi_dashboard_live_computed.sql
-- FIX: the Strategy KPI cards for event-derived financial metrics (di.* / parties.*)
-- were always null — v_kpi_dashboard only read metric_snapshots, and nobody logs
-- those by hand. Compute them LIVE from the event KPI views (completed events only)
-- and coalesce them into v_kpi_dashboard. Also adds 4 new high-value metrics.
-- FINANCIAL views — keep anon-revoked (E1 discipline / 015-019-022 guard).
-- =============================================================================
begin;

-- Live event-derived KPI values (+ new aggregate metrics). Completed events only.
create or replace view public.v_kpi_computed as
with di as (
  select k.* from public.v_kpi_dance_infusion k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
pt as (
  select k.* from public.v_kpi_parties k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
gk as (select * from public.v_guest_kpis limit 1)
select metric_key, value from (values
  ('di.raised_per_event',  (select round(avg(total_raised), 2) from di)),
  ('di.cost_to_raise',     (select round(avg(cost_to_raise_per_dollar), 2) from di)),
  ('di.attendance',        (select round(avg(total_attendance), 0) from di)),
  ('di.to_ms_total',       (select sum(net_pl) from di)),
  ('parties.net_pl',       (select round(avg(net_pl), 2) from pt)),
  ('parties.sell_through', (select round(avg(sell_through_pct), 1) from pt)),
  ('parties.net_pl_total', (select sum(net_pl) from pt)),
  ('audience.subscribers', (select count(*)::numeric from public.subscribers where status = 'subscribed')),
  ('guest.repeat_pct',     (select case when guests_with_attendance > 0 then round(100.0 * repeat_guests / guests_with_attendance, 1) end from gk))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

-- New KPI cards (targets). v_kpi_targets_current picks these up; values flow from v_kpi_computed.
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('di.to_ms_total',       'dance_infusion', 'To MS — total',     10000, 'gte', '$', current_date, true),
  ('parties.net_pl_total', 'parties',        'Net P&L — total',       0, 'gte', '$', current_date, true),
  ('audience.subscribers', 'audience',       'Mailing list',       1000, 'gte', '',  current_date, true),
  ('guest.repeat_pct',     'audience',       'Repeat attendees',     30, 'gte', '%', current_date, true)
on conflict do nothing;

-- Repoint the dashboard: computed value wins, else the latest manual snapshot.
create or replace view public.v_kpi_dashboard as
select t.metric_key, t.workstream, t.label,
  coalesce(c.value, l.value) as current_value,
  p.value as prior_value,
  t.target_value, t.comparison, t.unit,
  case when c.value is not null then current_date else l.captured_on end as as_of
from public.v_kpi_targets_current t
left join public.v_kpi_computed c on c.metric_key = t.metric_key
left join public.v_metric_latest l on l.metric_key = t.metric_key and l.series_id is null
left join public.v_metric_prior  p on p.metric_key = t.metric_key and p.series_id is null;
revoke select on public.v_kpi_dashboard from anon;

commit;

-- DOWN: drop view v_kpi_computed; recreate v_kpi_dashboard without the computed join;
--       delete the 4 new kpi_targets rows.
