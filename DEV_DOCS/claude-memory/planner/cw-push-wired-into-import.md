---
name: cw-push-wired-into-import
description: "2026-09-23: Run import now pushes Come With to the live site; the payload contract is pinned to the Edge Function and tested; UF_PUSH_DISABLED guards test harnesses"
metadata: 
  node_type: memory
  type: project
  originSessionId: 21c71d6f-845a-4732-ab58-5d2a0b40974b
  modified: 2026-09-23T22:54:36.711Z
---

**Until 2026-09-23 the finance import never pushed to the Come With site at all.** Jennifer's `Run import` updated `data.db` and stopped; the CW dashboard went stale for a month (45 rows / $1,746.93 of Aug–Sep spend). Both halves of the push existed but nothing joined them, and nothing called either: `uf_export_cw.py` only wrote `data/cw_push.json` to disk, `uf_push.py` had zero callers repo-wide.

**Now:** `run_all_imports_and_push()` in `scripts/uf_server.py` = import all three inboxes, then `push_cw()`. Both `/api/uf/import` handlers (`uf_server.py`, `src/web_server.py:~4200`) return `pushed` alongside `imported`, and `pushLine()` in `static/uf_dash_module.js` prints it in **every** branch of the import panel — including "no new files", because the push is a full true-up worth reporting even on an empty inbox.

**`push_cw()` never raises.** A failed push must not hide or roll back a good import; it comes back as a dict and a failed push turns the whole panel red. A silent push is the exact failure that caused this.

**The payload contract is now pinned, and it had drifted.** `uf_push.build_payload` was a placeholder whitelist that dropped `external_ref` — the key `ingest-finance` dedups and adopts on — so had the push ever run it would have duplicated the entire ledger instead of updating in place. `ROW_FIELDS`/`BUDGET_FIELDS` now mirror the `Row`/`BudgetLine` types in `supabase/functions/ingest-finance/index.ts`. `scripts/uf_push_test.py` (19 checks, sends nothing) parses those TS types and fails on real drift; when the Comewith repo isn't on the machine it reports SKIP, never a quiet pass.

**Deliberately not sent: `source`.** The site ignores it and it's the one field naming a personal feed ("simplifi_cw_mirror"). `funded_by` + `cash_source` already carry what the site needs.

**Send everything, every time.** The site reconciles on `external_ref`, so a full 241-row send is idempotent — verified live: first push 45 inserted / 196 updated, immediate re-push 0 inserted / 241 updated, 0 problems.

**`UF_PUSH_DISABLED=1` suppresses the push** and is set by `scripts/uf_import_smoke.py`. The HTTP smokes drive the real `/api/uf/import` against the live DB and delete what they insert — but **a push cannot be undone**, so a throwaway row that happened to route to Come With would land on the business site permanently. The skip is reported in the panel, never silent.

Related: [[comewith-repo-and-push-auth]], [[unified-finance-model]], [[pending-handoff-note]]
