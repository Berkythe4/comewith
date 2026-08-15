---
name: project-kpi-layer
description: "KPI/metrics layer (migration 015) applied to prod 2026-05-29 — schema only, dashboard wiring + backfill still pending"
metadata: 
  node_type: memory
  type: project
  originSessionId: d97501fd-d79e-460c-80ae-ea0889c23091
---

KPI/metrics layer applied to **prod** (`yaytdosxfhcqatmhctzk`) 2026-05-29 via the Supabase Management API query endpoint (PAT `SBP_PAT` in gitignored `.env`; prod ref `SBP_REF_PROD`). Tracked as `supabase/migrations/015_kpi.sql` with `015_kpi_handoff.md` alongside. **Schema only** — no events created, DI #1 backfill left commented (deferred by decision).

Created: `events.capacity` col; tables `content_series`, `metric_snapshots`, `kpi_targets` (all RLS `for all using (public.is_admin())`, seeded 3 series + 12 targets); views `v_kpi_event_financials` (thin reuse-wrapper over existing `v_event_summary` + events.capacity — NOT a duplicate), `v_kpi_parties`, `v_kpi_dance_infusion`, `v_kpi_dashboard`, `v_kpi_targets_current`, `v_metric_latest`, `v_metric_prior`. anon `select` REVOKED on the 4 financial KPI views + `v_event_summary` (decision E1); targets/metric views stay anon-readable per spec.

**Series contract (critical, must hold downstream):** `events.series` is free text, no CHECK. KPI views match EXACTLY — the future Log Event form MUST write `series = 'Come With Parties'` for parties and `series = 'Dance Infusion'` for DI, or those KPIs read empty. `'Come With Production'` is services, NOT parties.

Key schema facts confirmed on prod: `sponsorships` has NO `amount` col (use `cash_amount + in_kind_value`, exclude status='cancelled'); admin convention is helper `public.is_admin()` = role in ('master_admin','sub_admin') — there is NO 'admin' role (ties to [[project-phase-12-status]] role fact). Existing views have no `security_invoker`, so definer + anon grant = exposure unless revoked.

**Still pending (handoff §5):** dashboard Strategy section reading `v_kpi_dashboard` + per-event views; 3 entry forms (Log Event, Log Numbers, Edit Target — Edit Target inserts a NEW versioned row, never updates); DI #1 backfill once an event exists. Visual already designed in `ComeWith_Strategy_Dashboard.html`. Builds on [[project-phase-12-status]].
