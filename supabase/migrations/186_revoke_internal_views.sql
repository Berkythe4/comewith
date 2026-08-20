-- ============================================================
-- COME WITH — 186 three internal views were anon-readable
--
-- Found by the same sweep that turned up 185. These are not the ledger, but none
-- of them is anybody else's business, and nothing public reads them:
--
--   v_equipment_roi        equipment purchase prices, times used, revenue
--                          attributed per item. A financial view by any
--                          reasonable reading of decision E1.
--   v_mailing_list_health  subscribed / unsubscribed / bounced counts per
--                          segment. How big the list is and how well it is doing.
--   v_metric_prior         the previous value of every KPI. The internal
--                          scoreboard, with history.
--
-- Verified before revoking: no page under the site root reads any of them. Only
-- dashboard.html does, and that is behind admin auth.
--
-- NOT REVOKED, ON PURPOSE: v_kpi_targets_current. tools/visualizer.html reads it
-- ANONYMOUSLY - it has no sign-in at all - so revoking would break a working
-- internal tool without warning. It exposes targets (goals) rather than results,
-- which is the mildest of the four. Flagged for Keith to decide: either the
-- visualizer gets auth, or the view stays public deliberately. Left alone rather
-- than broken quietly.
-- ============================================================
begin;

revoke select on public.v_equipment_roi       from anon;
revoke select on public.v_mailing_list_health from anon;
revoke select on public.v_metric_prior        from anon;

commit;

-- DOWN: grant select on those three back to anon (nothing needs it).
