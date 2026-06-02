-- =============================================================================
-- 027_measurement_layer.sql  —  Phase D: measurement + definitions (additive)
-- Spec §2, §3, §6 Phase D. NOT APPLIED — review before apply. Push held.
--
-- metric_definitions (KPIs/formulas as data); v_data_points (uniform stream:
-- snapshots live + derived materialized); nightly pg_cron refresh of the derived
-- MV. Tier-2 formula evaluation runs client-side in the visualizer (Phase E) —
-- no SQL eval, no arbitrary code (Q5).
--
-- FINANCIAL: mv_event_data_points + v_data_points expose per-event money. Revoked
-- from anon here; MUST also be revoked from `authenticated` before any external
-- login (ROADMAP blocker / BUILD_LOG §2 — same gate as the 5 KPI views).
-- =============================================================================

-- ---- Layer 3: metric_definitions ----
create table public.metric_definitions (
  id           uuid primary key default gen_random_uuid(),
  metric_key   text not null unique,
  label        text not null,
  unit         text,                 -- $ | % | hrs | count | ''
  category     text,                 -- content | parties | audience | dance_infusion | ...
  comparison   text not null default 'gte' check (comparison in ('gte','lte')),
  source_type  text not null check (source_type in ('snapshot','derived','formula')),
  source_ref   text,                 -- snapshot: snapshot metric_key; derived: view.column; formula: expression over metric_keys
  active       boolean not null default true,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);
create trigger set_updated_at before update on public.metric_definitions
  for each row execute function public.handle_updated_at();
alter table public.metric_definitions enable row level security;
create policy "Admins can manage metric definitions" on public.metric_definitions for all using (public.is_admin());

-- ---- Layer 2: derived per-event data points (MATERIALIZED — scale) ----
create materialized view public.mv_event_data_points as
select d.metric_key,
       'event'::text                              as subject_type,
       vs.event_id                                as subject_id,
       vs.event_date                              as captured_at,
       d.value                                    as value,
       jsonb_build_object('event_type', e.type, 'series', vs.series, 'venue_id', vs.venue_id) as dims
  from public.v_event_summary vs
  join public.events e on e.id = vs.event_id
  cross join lateral (values
     ('event.net_pl',          vs.net),
     ('event.ticket_revenue',  vs.ticket_revenue),
     ('event.other_income',    vs.revenue),
     ('event.donations',       vs.third_party_donations),
     ('event.sponsor_cash',    vs.sponsor_cash),
     ('event.sponsor_in_kind', vs.sponsor_in_kind),
     ('event.sponsor_count',   vs.sponsor_count::numeric),
     ('event.tickets_sold',    vs.tickets_sold::numeric),
     ('event.attendance',      vs.total_attendance::numeric),
     ('event.gross_revenue',   vs.ticket_revenue + vs.revenue + vs.third_party_donations + vs.sponsor_cash),
     ('event.total_raised',    vs.ticket_revenue + vs.revenue + vs.third_party_donations + vs.sponsor_cash + vs.sponsor_in_kind)
  ) as d(metric_key, value)
 where d.value is not null;

create index idx_mv_edp_key_date on public.mv_event_data_points(metric_key, captured_at);
create index idx_mv_edp_subject on public.mv_event_data_points(subject_type, subject_id);

-- ---- Layer 2: the uniform stream (snapshots LIVE + derived from the MV) ----
create or replace view public.v_data_points as
  select s.metric_key,
         case when s.series_id is not null then 'content_series' else 'global' end as subject_type,
         s.series_id  as subject_id,
         s.captured_on as captured_at,
         s.value,
         '{}'::jsonb  as dims
    from public.metric_snapshots s
  union all
  select metric_key, subject_type, subject_id, captured_at, value, dims
    from public.mv_event_data_points;

-- Financial — lock from anon (E1). NOTE: also lock from authenticated before logins.
revoke select on public.mv_event_data_points from anon;
revoke select on public.v_data_points        from anon;

-- ---- nightly refresh of the derived MV (pg_cron from 014) ----
-- Refresh at 03:00. (Non-concurrent: brief lock at 3am is fine; concurrent would
-- require a unique index the unpivoted rows don't naturally have.)
select cron.schedule('refresh-mv-event-data-points', '0 3 * * *',
  $$refresh materialized view public.mv_event_data_points$$);

-- ---- seed metric_definitions (derived event metrics + a Tier-2 formula example) ----
insert into public.metric_definitions (metric_key, label, unit, category, comparison, source_type, source_ref) values
  ('event.net_pl',          'Net P&L',            '$',     'parties',        'gte', 'derived', 'v_event_summary.net'),
  ('event.gross_revenue',   'Gross revenue',      '$',     'parties',        'gte', 'derived', 'computed'),
  ('event.total_raised',    'Total raised',       '$',     'dance_infusion', 'gte', 'derived', 'computed'),
  ('event.ticket_revenue',  'Ticket revenue',     '$',     'parties',        'gte', 'derived', 'v_event_summary.ticket_revenue'),
  ('event.donations',       'Donations',          '$',     'dance_infusion', 'gte', 'derived', 'v_event_summary.third_party_donations'),
  ('event.sponsor_count',   'Sponsors',           'count', 'parties',        'gte', 'derived', 'v_event_summary.sponsor_count'),
  ('event.tickets_sold',    'Tickets sold',       'count', 'parties',        'gte', 'derived', 'v_event_summary.tickets_sold'),
  ('event.attendance',      'Attendance',         'count', 'parties',        'gte', 'derived', 'events.total_attendance'),
  ('instagram.followers',   'Instagram followers','count', 'audience',       'gte', 'snapshot','instagram.followers'),
  -- Tier-2 formula: composes other metric_keys (+ - * / and parens), evaluated client-side.
  ('event.revenue_per_attendee', 'Revenue / attendee', '$', 'parties',       'gte', 'formula', 'event.gross_revenue / event.attendance')
on conflict (metric_key) do nothing;

-- Grants: 013 default privileges; metric_definitions admin-only; v_data_points +
-- mv anon-revoked (financial). No anon grants.

-- DOWN: select cron.unschedule('refresh-mv-event-data-points');
--       drop view v_data_points; drop materialized view mv_event_data_points;
--       drop table metric_definitions;
