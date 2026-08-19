---
name: preexisting-frontend-test-failures
description: Stage 26 nav redesign broke 8 flat-nav frontend tests; fixed 2026-05-28. Only remaining expected full-suite failure is the date-relative refiner flake.
metadata: 
  node_type: memory
  type: project
  originSessionId: 23962f70-e541-4462-96aa-41148968baf0
---

A full `pytest tests/` run on 2026-05-28 showed 9 failures. **8 of them were stale flat-nav assertions and are now fixed; 1 remains as a known flake.**

**Fixed 2026-05-28** — `tests/test_stage_05_5_visuals_frontend.py` and `tests/test_stage_07_finance_frontend.py` asserted `data-route="/finance"` / `data-route="/visuals"` in page topbars. Stage 26's nav redesign moved Finance/Visuals into the **Tools ▾ / Dev ▾ dropdowns**, which use `href=` + `data-surface="finance"` / `data-surface="visuals"` (no `data-route`; that attribute now only lives on the primary Chat/Calendar/Whiteboard tier). The tests were updated to assert the dropdown structure.

**Still failing (known flake)** — `tests/test_polish.py::test_chat_plan_refiner_respects_pause` is date-relative (noted in `SESSION_RESUME.md`). This is the only expected full-suite failure now.

**Why:** documents that the post-Stage-26 nav uses `data-surface=` in dropdowns, not `data-route=`, so future nav/test work matches the real structure — and that a clean full suite means 0 failures except the one flake.

**How to apply:** if a full-suite run shows more than the single refiner flake, it's likely a real regression. Relates to [[bug-sweep-2026-05-28]].
