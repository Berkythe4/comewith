---
name: project_impact_report_supabase
description: DI impact report wired to Supabase (migration 067) — BUILT but NOT applied to prod; dashboard editor + publish toggle + homepage button
metadata: 
  node_type: memory
  type: project
  originSessionId: 59fd8c20-e65f-4382-bad4-d875b13df3be
---

Dance Infusion impact report + public audit moved from static local JSON to a
Supabase-backed, dashboard-editable, toggle-published feature (build 2026-06-28).

**Migration `067_impact_report.sql` is WRITTEN but NOT YET APPLIED to prod** — the
prod `SBP_PAT` in `.env` returned 401 (expired) during the build, and prod applies
are user-initiated. To finish go-live: refresh SBP_PAT, then run
`scripts/apply_067_impact_report.py` (applies 067 + seeds the DI#2 jsonb from the
local JSON + sets internal KPI `di.cost_to_raise`=0.50). Until applied, the public
pages fall back to local JSON **only on localhost**; on the live site they show
"not published yet".

- 067 adds `events.impact_report jsonb` + `events.impact_report_public boolean`
  (the publish toggle) + anon view `v_public_impact_report` (returns a row only when
  the toggle is on). No financial views touched — audit figures are curated values in
  the jsonb, not the anon-revoked v_kpi_*/v_event_summary.
- **Dashboard editor**: event-hub header button "Impact report" (Dance Infusion events
  only) → `openImpactReportEditor`/`submitImpactReport` in dashboard.html. Edits text +
  2 photos (hero/inline, event-photos bucket, HEIC pipeline) + the toggle; preserves
  seeded numbers on save.
- **Public pages** (`events/dance-infusion/di-02-2026-05/reports/impact-report.html` +
  `public-audit.html`): read `v_public_impact_report` first, fallback to local JSON on
  localhost only. Old `/staging/` guard removed — the DB toggle is now the gate.
- **Homepage** `#di` section shows a "Read the #2 Impact Report →" button only when a
  report is published (`loadImpactReport()` in index.html).
- Copy decisions baked in: attendance 42→**117 on the floor**, DI#1 sponsors **0**,
  **reach section removed**, public goal **50% to mission** (was 60). See
  [[project_historical_events_backfill]] for DI money, [[project_public_events]] for the
  anon-view pattern, [[feedback_prod_migration_apply]] for the apply mechanism.
- Nothing pushed/deployed yet — all local working-tree changes.
