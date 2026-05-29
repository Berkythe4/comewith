-- ============================================================
-- COME WITH — KPI / METRICS LAYER
-- Additive migration. Does not alter or drop existing tables.
-- Apply via Claude Code against project yaytdosxfhcqatmhctzk.
--
-- Reconciled against the live prod schema 2026-05-29: all [VERIFY]
-- refs resolved (expenses/income/ticketing/donations confirmed;
-- sponsorships has no `amount` — uses cash_amount+in_kind_value via
-- v_event_summary reuse); RLS uses public.is_admin(); series matched
-- exactly to 'Come With Parties' / 'Dance Infusion'; anon revoked on
-- financial views. Schema only — no events created, backfill left commented.
-- ============================================================

begin;

-- ------------------------------------------------------------
-- 1. EVENT CAPACITY  (needed for sell-through %)
--    events has total_attendance + bar_minimum but no capacity.
-- ------------------------------------------------------------
alter table public.events
  add column if not exists capacity integer;

-- ------------------------------------------------------------
-- 2. CONTENT SERIES  (YouTube series identity — thumbnail-level)
--    e.g. Backyard Sessions, Live at Dance Infusion, Come With Sets
-- ------------------------------------------------------------
create table if not exists public.content_series (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  slug        text unique,
  kind        text,                 -- 'showcase' | 'event_recording' | 'dj_set'
  active      boolean default true,
  created_at  timestamptz default now()
);

-- ------------------------------------------------------------
-- 3. METRIC SNAPSHOTS  (periodic readings with no transactional home)
--    Instagram followers/engagement, YouTube subs/watch time/views,
--    per-series views. One row = one reading on one day.
--    source = 'manual' now; 'youtube_api' / 'instagram_api' later.
-- ------------------------------------------------------------
create table if not exists public.metric_snapshots (
  id          uuid primary key default gen_random_uuid(),
  metric_key  text not null,        -- e.g. 'instagram.followers', 'youtube.subscribers'
  value       numeric not null,
  captured_on date not null default current_date,
  series_id   uuid references public.content_series(id),  -- nullable; set for per-series content metrics
  source      text default 'manual',
  notes       text,
  created_by  uuid,
  created_at  timestamptz default now()
);

-- one reading per metric (per series) per day
create unique index if not exists metric_snapshots_unique
  on public.metric_snapshots (
    metric_key,
    captured_on,
    coalesce(series_id, '00000000-0000-0000-0000-000000000000'::uuid)
  );

create index if not exists metric_snapshots_key_date
  on public.metric_snapshots (metric_key, captured_on desc);

-- ------------------------------------------------------------
-- 4. KPI TARGETS  (editable from the dashboard — never edited in SQL)
--    Versioned by effective_date: changing a target inserts a new
--    row; the "current" target is the latest active row per metric.
--    comparison: 'gte' = higher is better, 'lte' = lower is better
--    (e.g. cost-to-raise-a-dollar uses 'lte').
-- ------------------------------------------------------------
create table if not exists public.kpi_targets (
  id             uuid primary key default gen_random_uuid(),
  metric_key     text not null,
  workstream     text not null,     -- 'content' | 'parties' | 'audience' | 'dance_infusion'
  label          text not null,     -- shown on the dashboard card
  target_value   numeric not null,
  comparison     text not null default 'gte',
  unit           text,              -- '%' | '$' | 'hrs' | '' etc.
  effective_date date not null default current_date,
  active         boolean default true,
  updated_by     uuid,
  updated_at     timestamptz default now()
);

create index if not exists kpi_targets_key_date
  on public.kpi_targets (metric_key, effective_date desc);

-- current target per metric
create or replace view public.v_kpi_targets_current as
select distinct on (metric_key)
  metric_key, workstream, label, target_value, comparison, unit, effective_date
from public.kpi_targets
where active
order by metric_key, effective_date desc;

-- ------------------------------------------------------------
-- 5. LATEST METRIC VALUE + PRIOR READING (for trend arrows)
-- ------------------------------------------------------------
create or replace view public.v_metric_latest as
select distinct on (metric_key, series_id)
  metric_key, series_id, value, captured_on
from public.metric_snapshots
order by metric_key, series_id, captured_on desc;

-- value as of the most recent reading strictly before the latest
create or replace view public.v_metric_prior as
select distinct on (metric_key, series_id)
  s.metric_key, s.series_id, s.value, s.captured_on
from public.metric_snapshots s
join public.v_metric_latest l
  on l.metric_key = s.metric_key
 and coalesce(l.series_id,'00000000-0000-0000-0000-000000000000'::uuid)
   = coalesce(s.series_id,'00000000-0000-0000-0000-000000000000'::uuid)
 and s.captured_on < l.captured_on
order by s.metric_key, s.series_id, s.captured_on desc;

-- ------------------------------------------------------------
-- 6. EVENT FINANCIAL ROLLUP  (foundation for party + DI KPIs)
--    Reuses the existing public.v_event_summary (single source of
--    truth for per-event financials) and only adds events.capacity.
--    v_event_summary already filters income/expenses on deleted_at,
--    excludes cancelled sponsorships, and computes sponsor_cash as
--    (cash_amount + in_kind_value) — sponsorships has no `amount`.
--    Downstream column names kept: total_income / total_expenses /
--    donations / sponsor_value.
-- ------------------------------------------------------------
create or replace view public.v_kpi_event_financials as
select
  s.event_id,
  s.name,
  s.series,
  s.event_date,
  e.capacity,
  s.total_attendance,
  s.tickets_sold,
  s.ticket_revenue,
  s.revenue               as total_income,
  s.expenses              as total_expenses,
  s.third_party_donations as donations,
  s.sponsor_cash          as sponsor_value
from public.v_event_summary s
join public.events e on e.id = s.event_id;

-- ------------------------------------------------------------
-- 7. WORKSTREAM KPI VIEWS
--    events.series is the internal category (free text, no CHECK).
--    Canonical strings the Log Event form MUST write:
--      'Come With Parties'  → party events
--      'Dance Infusion'     → DI events
--    Exact match (not ilike) so e.g. 'Come With Production' is excluded.
-- ------------------------------------------------------------

-- Come With Parties — sell-through + net P&L per event
create or replace view public.v_kpi_parties as
select
  event_id, name, event_date, capacity, tickets_sold,
  case when capacity > 0
       then round(tickets_sold::numeric / capacity * 100, 1) end as sell_through_pct,
  (total_income - total_expenses)                                 as net_pl
from public.v_kpi_event_financials
where series = 'Come With Parties';

-- Dance Infusion — cost to raise a dollar + total raised + attendance
-- total_raised = ticket_revenue + donations + sponsor_value (gross)
-- cost_to_raise = total_expenses / total_raised  (lower is better)
create or replace view public.v_kpi_dance_infusion as
select
  event_id, name, event_date, total_attendance,
  (ticket_revenue + donations + sponsor_value)                    as total_raised,
  case when (ticket_revenue + donations + sponsor_value) > 0
       then round(total_expenses
                  / nullif(ticket_revenue + donations + sponsor_value,0), 2)
  end                                                             as cost_to_raise_per_dollar
from public.v_kpi_event_financials
where series = 'Dance Infusion';

-- ------------------------------------------------------------
-- 8. UNIFIED DASHBOARD FEED
--    One row per tracked metric: current value, prior value,
--    target, comparison, unit, workstream. The dashboard reads
--    THIS for the snapshot-style metrics (followers, subs, etc.).
--    Event-derived metrics (sell-through, P&L, cost-to-raise)
--    read the per-event views above for the latest event.
-- ------------------------------------------------------------
create or replace view public.v_kpi_dashboard as
select
  t.metric_key,
  t.workstream,
  t.label,
  l.value          as current_value,
  p.value          as prior_value,
  t.target_value,
  t.comparison,
  t.unit,
  l.captured_on    as as_of
from public.v_kpi_targets_current t
left join public.v_metric_latest l on l.metric_key = t.metric_key and l.series_id is null
left join public.v_metric_prior  p on p.metric_key = t.metric_key and p.series_id is null;

-- ------------------------------------------------------------
-- 9. SEED TARGETS  (the starting targets from the dashboard;
--    Keith & Liz edit these from the UI afterward)
-- ------------------------------------------------------------
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit) values
  ('youtube.subscribers',        'content',        'Subscribers',          2000, 'gte', ''),
  ('youtube.watch_hours',        'content',        'Watch time / mo',       750, 'gte', 'hrs'),
  ('youtube.avg_views',          'content',        'Avg views / upload',   4000, 'gte', ''),
  ('parties.presale_velocity',   'parties',        'Pre-sale velocity',      70, 'gte', '%'),
  ('parties.sell_through',       'parties',        'Sell-through',          100, 'gte', '%'),
  ('parties.net_pl',             'parties',        'Net P&L / event',         0, 'gte', '$'),
  ('instagram.followers',        'audience',       'Instagram followers',  5000, 'gte', ''),
  ('instagram.saves_shares',     'audience',       'Saves + shares / post', 150, 'gte', ''),
  ('audience.follower_ticket',   'audience',       'Follower → ticket',      10, 'gte', '%'),
  ('di.cost_to_raise',           'dance_infusion', 'Cost to raise $1',     0.25, 'lte', '$'),
  ('di.raised_per_event',        'dance_infusion', 'Raised / event',       6000, 'gte', '$'),
  ('di.attendance',              'dance_infusion', 'Attendance',            250, 'gte', '')
on conflict do nothing;

-- ------------------------------------------------------------
-- 10. SEED CONTENT SERIES
-- ------------------------------------------------------------
insert into public.content_series (name, slug, kind) values
  ('Backyard Sessions',     'backyard-sessions',     'showcase'),
  ('Live at Dance Infusion','live-at-dance-infusion','event_recording'),
  ('Come With Sets',        'come-with-sets',        'dj_set')
on conflict (slug) do nothing;

-- ------------------------------------------------------------
-- 11. RLS  — admin-only, matching the project convention:
--    every operational table uses  for all using (public.is_admin()).
--    is_admin() = role in ('master_admin','sub_admin'). There is NO
--    'admin' role. For an ALL policy with only USING, the WITH CHECK
--    defaults to the USING expression, so admin inserts pass too.
-- ------------------------------------------------------------
alter table public.content_series   enable row level security;
alter table public.metric_snapshots enable row level security;
alter table public.kpi_targets       enable row level security;

create policy "Admins can manage content_series"
  on public.content_series   for all using (public.is_admin());
create policy "Admins can manage metric_snapshots"
  on public.metric_snapshots for all using (public.is_admin());
create policy "Admins can manage kpi_targets"
  on public.kpi_targets       for all using (public.is_admin());

-- ------------------------------------------------------------
-- 12. GRANTS  — idempotent, mirrors 013_grants so the new tables and
--    views are reachable for policy evaluation by the app roles.
--    Row-level access is still gated by the RLS policies above.
-- ------------------------------------------------------------
grant usage on schema public to anon, authenticated, service_role;
grant all on all tables    in schema public to anon, authenticated, service_role;
grant all on all sequences in schema public to anon, authenticated, service_role;
grant all on all functions in schema public to anon, authenticated, service_role;

-- ------------------------------------------------------------
-- 13. REVOKE anon on financial views (decision E1). These views run
--    with definer rights and would otherwise let anon bypass the RLS
--    on the underlying financial tables. The dashboard reads them as
--    an authenticated admin and is unaffected.
-- ------------------------------------------------------------
revoke select on public.v_event_summary        from anon;
revoke select on public.v_kpi_event_financials from anon;
revoke select on public.v_kpi_parties          from anon;
revoke select on public.v_kpi_dance_infusion   from anon;
revoke select on public.v_kpi_dashboard        from anon;

commit;

-- ============================================================
-- DI #1 BACKFILL TEMPLATE  (run after confirming column names)
-- Fill the real numbers, then apply. Establishes the baseline
-- so "growth per event" has something to grow from.
-- ============================================================
-- update public.events
--   set capacity = :di1_capacity, total_attendance = :di1_attendance
--   where slug = 'dance-infusion-1';   -- [VERIFY slug]
--
-- insert into public.metric_snapshots (metric_key, value, captured_on, source, notes) values
--   ('instagram.followers', :ig_at_di1, '2025-09-06', 'manual', 'baseline at DI #1'),
--   ('youtube.subscribers', :yt_at_di1, '2025-09-06', 'manual', 'baseline at DI #1');
-- (ticket / donation / expense rows for DI #1 go into ticketing /
--  third_party_donations / expenses with the DI #1 event_id.)
