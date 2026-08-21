-- ============================================================
-- COME WITH — 187 the last knowingly-public internal view
--
-- 186 left v_kpi_targets_current granted to anon on a stated premise:
-- "tools/visualizer.html reads it ANONYMOUSLY - it has no sign-in at all".
--
-- That premise was wrong. tools/visualizer.html line 7 loads /staging/guard.js
-- and line 60 imports its `sb` client from it — the same admin gate every
-- /staging/ page uses. With no session the guard redirects to /dashboard.html.
--
-- And the tool could never have worked as anon anyway. Asked as the public
-- (2026-08-21, real publishable key):
--
--   metric_definitions      200 []    <- RLS returns nothing, so the metric
--                                        picker is empty and there is nothing
--                                        to chart
--   v_data_points           401       <- already revoked; no series to draw
--   v_kpi_targets_current   200 rows  <- the only one that answered
--
-- So the anon grant was not holding a working tool up. It was handing the
-- public every KPI target we have set — the goals, per workstream, with their
-- effective dates. Targets rather than results, which is why 186 rated it the
-- mildest of the four, but it is nobody else's business either.
--
-- `authenticated` keeps SELECT (verified before writing this), so the
-- visualizer and the dashboard are unaffected — they read with a session.
--
-- After this, nothing in public is anon-readable except the public site feed.
-- ============================================================
begin;

revoke select on public.v_kpi_targets_current from anon;

commit;

-- DOWN: grant select on public.v_kpi_targets_current to anon;
--       (nothing needs it — restoring this re-opens the leak.)
