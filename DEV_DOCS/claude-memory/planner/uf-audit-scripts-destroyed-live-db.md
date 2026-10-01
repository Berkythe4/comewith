---
name: uf-audit-scripts-destroyed-live-db
description: Three scripts/uf_* audits used to re-migrate the LIVE data.db and wipe every import since June — guarded 2026-08-21
metadata: 
  node_type: memory
  type: project
  originSessionId: 2dcea4f4-a99e-473f-8561-1294e0a958b7
  modified: 2026-08-21T20:44:31.126Z
---

2026-08-21: running `scripts/uf_server_test.py` against the live DB destroyed 439
Personal and 114 Come With transactions. `uf_final_e2e.py` and `uf_phase2_audit.py`
did the same thing. All three "reverted to a clean state" by re-running
`uf_phase1_migrate.py`, which DELETEs every `uf_*` table and rebuilds it from the
golden workbooks in `data/`. That was safe in June, when data.db was nothing but
the migration's output. It has not been safe since the first import.

**Why:** the golden Excel files are the migration's INPUT and are frozen. Every
Simplifi / PayPal / Bluevine import, every hand-entered charge, and every income
line lives ONLY in data.db, so a re-migration silently discards all of it. Nothing
warned; the scripts printed "reverted to clean canonical state" and exited 0.

**How to apply:** recovered from `data.db.bak_before_income_center_20260821_152403`
(take a timestamped backup before ANY uf_* work — the damaged copy is kept as
`data.db.damaged_by_uf_server_test_*`). The fix is now in three places:
`uf_phase1_migrate` REFUSES to run when the DB holds more transactions than the
Excel can restore (`--force` overrides, and it always backs up first), and it
accepts a target DB path; `uf_phase1_audit` / `uf_phase2_audit` / `uf_final_e2e`
migrate a throwaway DB under `data/_sandbox/` and redirect `SYNC.MIRROR` there, so
they never open data.db; `uf_server_test` puts back only the one budget cell and
one input it changes. Before trusting any uf_* audit that claims a clean tie-out,
check whether it is reading the sandbox or the real books — phase 1/2 and the e2e
are sandbox-only by design now, and only `uf_phase3_audit` reconciles the LIVE DB
against the regenerated mirror. Related: [[unified-finance-model]],
[[finance-income-control-center]].
