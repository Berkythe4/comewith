-- =============================================================================
-- 014_cron.sql
-- pg_cron schedules for background automation.
-- pg_cron is already enabled in 001_extensions.sql.
--
-- Two jobs:
--   1. refresh-materialized-views — nightly MV refresh (03:00 UTC)
--   2. audit-log-retention        — daily delete of audit_log rows
--                                   older than 365 days (04:00 UTC)
--
-- A third job (scheduled-campaign-sends) is intentionally deferred:
-- send-campaign requires an admin JWT, which pg_cron can't construct
-- without storing the service_role key in vault. Documented in
-- project_phase_10_status memory; can be added once the auth design
-- is decided (cron-secret header + send-campaign accepting it, OR
-- vault + service_role token in pg_net headers).
--
-- Safe to re-run: cron.schedule returns the same job id if a job
-- with the same name already exists. automation_jobs is upserted.
-- =============================================================================

-- 1. Refresh materialized views nightly at 03:00 UTC.
--    Uses CONCURRENTLY (each MV has a unique index from 011_views.sql)
--    so dashboard queries aren't blocked during refresh.
select cron.schedule(
  'refresh-materialized-views',
  '0 3 * * *',
  $$
    refresh materialized view concurrently public.mv_cross_event_kpis;
    refresh materialized view concurrently public.mv_repeat_sponsors;
    refresh materialized view concurrently public.mv_top_artists;
  $$
);

-- 2. Audit log retention. Keep 365 days; delete older rows daily at 04:00 UTC.
--    Adjust the interval here if compliance requires a longer or shorter window.
select cron.schedule(
  'audit-log-retention',
  '0 4 * * *',
  $$
    delete from public.audit_log
    where occurred_at < now() - interval '365 days';
  $$
);

-- Register jobs in automation_jobs so they show up in the dashboard.
insert into public.automation_jobs (name, description, cron_expression, edge_function, enabled)
values
  (
    'refresh-materialized-views',
    'Nightly refresh of mv_cross_event_kpis, mv_repeat_sponsors, mv_top_artists',
    '0 3 * * *',
    'inline-sql',
    true
  ),
  (
    'audit-log-retention',
    'Delete audit_log rows older than 365 days',
    '0 4 * * *',
    'inline-sql',
    true
  )
on conflict (name) do update set
  description = excluded.description,
  cron_expression = excluded.cron_expression,
  enabled = excluded.enabled;
