-- ============================================================
-- COME WITH — 019 RE-REVOKE anon ON KPI FINANCIAL VIEWS  (corrective)
--
-- Regression fix: migrations 016 and 017 each included a broad
-- `grant all on all tables in schema public to anon`, which re-granted
-- anon SELECT on the financial KPI views that 015 had revoked (E1).
-- v_event_summary was already re-revoked in 018; the four KPI views
-- were left exposed. Re-revoke them. Pure REVOKE — the dashboard reads
-- these as an authenticated admin and is unaffected.
--
-- Lesson: do NOT use broad `grant all on all tables to anon` in future
-- migrations — 013's ALTER DEFAULT PRIVILEGES already covers new tables,
-- and the broad grant silently re-exposes revoked views.
-- ============================================================
begin;

revoke select on public.v_kpi_event_financials from anon;
revoke select on public.v_kpi_parties          from anon;
revoke select on public.v_kpi_dance_infusion   from anon;
revoke select on public.v_kpi_dashboard        from anon;
revoke select on public.v_event_summary        from anon;  -- idempotent re-assert

commit;
