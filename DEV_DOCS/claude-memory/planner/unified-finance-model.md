---
name: unified-finance-model
description: "uf_* unified finance schema in data.db — Jennifer canonical, Excel synced mirror, CW→Personal income link"
metadata: 
  node_type: memory
  type: project
  originSessionId: 1bb87365-2460-4e57-ab2a-d85b33c7e8e5
---

2026-06-19 overnight build: made Jennifer's DB canonical for BOTH budgets. Added 7 ADDITIVE prefixed tables to the LIVE app DB `data.db` (NOT the stale `data/planner.db`): `uf_envelopes`, `uf_transactions`, `uf_budgets`, `uf_cw_earnings`, `uf_runway_inputs`, `uf_review_queue`, `uf_rules` + views `uf_v_keith_income`/`uf_v_personal_actual`/`uf_v_cw_actual`. (Reversal of the original "never touch Jennifer's DB" rule — the user explicitly authorized it.) `rules` as a bare name COLLIDES with an existing table → that's why the `uf_` prefix.

**Why:** unify Personal.xlsx + Come With budgets under one canonical store with a Jennifer dashboard; Excel becomes a regenerated mirror.

**How to apply:** Migrated from the VALIDATED Excel (not the 6,669-row legacy `finance_transactions`). Entity explicit on every row (Personal | Come With); signed convention (expense negative, income positive). CW→Personal link = payroll model: only KEITH's ACTUAL CW earnings (`uf_v_keith_income`, from `uf_cw_earnings` person=Keith) post to Personal income; Martin's share + CW 15% net stay in CW. CW earnings derived from CW Budget&P&L DJ-gig pay (cost==85%×rev signature) split 50/50 (3-Year Budget r36). Dedup hash = `sha256(entity|date|amount|desc|occurrence)` (lossless for duplicate charges, idempotent). Scripts in `scripts/uf_*.py`: `uf_model` (compute/runway), `uf_phase1_migrate`, `uf_ingest` (drop→route→review queue→remember-rule), `uf_excel_sync` (regenerate mirror to `data/mirror/` + recon, golden files untouched), `uf_dashboard_build` (standalone dashboard `data/uf_dashboard/index.html`, reads only uf_*), + `uf_phase{1,2,3}_audit` and `uf_final_e2e`. All audits GREEN. CW P&L "Mo" column is bare month abbrevs = Year-1 2026 (MONTHMAP). Deferred for human review: live `web_server.py` dashboard route wiring (kept additive/standalone to avoid risking the running app); CW reserve $5k placeholder. See `FINAL_REPORT.md`, `BUILD_LOG.md`, `CLEANUP_MANIFEST.md`. Related: [[personal-xlsx-simplifi-import]].
