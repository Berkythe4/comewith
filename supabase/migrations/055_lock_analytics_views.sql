-- =============================================================================
-- 055_lock_analytics_views.sql  (closes the residual financial-view blocker)
-- v_budget_variance / v_data_points expose budget + financial data and still
-- grant SELECT to `authenticated` (any staff login could read them via REST).
-- mv_event_data_points is a MATERIALIZED view (RLS/security_invoker do not apply
-- to it), so the only safe lock is to revoke role access entirely.
--
-- None of the three are used by dashboard.html. Master reaches financials through
-- the gated event views (043); these stay service-role / owner only (the nightly
-- refresh runs as the table owner, unaffected). Revoke ALL from anon + auth.
-- =============================================================================
begin;
revoke all on public.v_budget_variance   from anon, authenticated;
revoke all on public.v_data_points        from anon, authenticated;
revoke all on public.mv_event_data_points from anon, authenticated;
commit;
-- POST: a staff (authenticated) REST GET on any of the three -> 401/permission denied.
-- ROLLBACK: grant select on <view> to authenticated;  (do NOT re-grant to anon.)
