---
name: project-phase-3-status
description: Phase 3 (admin dashboard read-only) closed 2026-05-28. dashboard-v2.html reads all 7 admin tables from Supabase staging via supabase-js. Pattern is in place for Phase 4 writes.
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 3 closed 2026-05-28 (same day as Phase 2). Six sprints, all
committed individually (commits 6d6d0c2 through 7424809). Final
commit: `7424809 Phase 3 sprint 6: polish + error handling + boot screen`.

## What ships
- `dashboard-v2.html` at repo root (parallel to existing dashboard.html,
  doesn't touch the live prod surface)
- Magic-link login via `supabase-js@2` from jsdelivr CDN
- 7 read-only tabs: inquiries, agreements, clients, income, expenses,
  equipment, events
- Per-tab loader pattern: `queries[tab]` + `renderers[tab]` —
  Phase 4 writes plug in alongside without refactoring

## Out-of-band staging changes (not in git)
- Supabase Auth `uri_allow_list` extended to include
  `http://localhost:8765/**` and `http://127.0.0.1:8765/**`
  (port 3000 is taken on the dev machine by another app)
- Supabase Site URL still `http://localhost:3000` (placeholder)

## Verified counts at Phase 3 close
- inquiries: 1 (leftover phase2-test row from Phase 2 debugging,
  never cleaned up because dashboard warned on DELETE)
- agreements: 1 (leftover phase2-smoketest)
- clients: 2 (Jared from Phase 1 + Test Customer from Phase 2)
- income: 6 (Phase 1 import)
- expenses: 62 (Phase 1 import)
- equipment_inventory: 13 (Phase 1 import)
- events: 0

## Open for Phase 4
- Write flows: create/edit inquiries, transition agreement status,
  add income/expense, equipment CRUD. The locked Phase 4 scope from
  the roadmap is "Berky stops using Sheets day-to-day."
- Decide modal/inline-edit UX before building
- Add Phase 4 buttons and forms; reuse the existing data-table CSS
- The phase2-test inquiry + agreement are still in staging; can be
  cleaned with `delete from inquiries where source = 'phase2-test'`
  and `delete from agreements where notes = 'RLS smoke test agreement'`
  whenever convenient

## How to apply
- The pattern in dashboard-v2.html (queries object + renderers object
  + loadTab function + per-tab error/empty states) is the template
  for any new tab or write-capable tab. Stick to it for consistency.
- `db.py` works for any SQL operation on staging; use it instead of
  the SQL Editor unless the user explicitly wants paste-and-confirm.
- The local dev server lives at port 8765, not 3000.

Related: [[project-phase-2-status]], [[project-anon-rls-sql-editor]]
