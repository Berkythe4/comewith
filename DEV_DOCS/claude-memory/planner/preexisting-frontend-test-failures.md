---
name: preexisting-frontend-test-failures
description: What a clean `pytest tests/` looks like — 3 known failures as of 2026-08-21; anything more is a real regression.
metadata: 
  node_type: memory
  type: project
  originSessionId: 23962f70-e541-4462-96aa-41148968baf0
---

A full `pytest tests/` run on 2026-05-28 showed 9 failures. **8 of them were stale flat-nav assertions and are now fixed; 1 remains as a known flake.**

**Fixed 2026-05-28** — `tests/test_stage_05_5_visuals_frontend.py` and `tests/test_stage_07_finance_frontend.py` asserted `data-route="/finance"` / `data-route="/visuals"` in page topbars. Stage 26's nav redesign moved Finance/Visuals into the **Tools ▾ / Dev ▾ dropdowns**, which use `href=` + `data-surface="finance"` / `data-surface="visuals"` (no `data-route`; that attribute now only lives on the primary Chat/Calendar/Whiteboard tier). The tests were updated to assert the dropdown structure.

**Baseline as of 2026-08-21 — a clean `pytest tests/` is exactly THREE failures:**
1. `test_polish.py::test_chat_plan_refiner_respects_pause` — date-relative flake (also in `SESSION_RESUME.md`).
2. `test_suggestions_and_decisions.py::test_blocking_check_response` — expects a paused "CWF" parent to surface; depends on the live `config/parents.yaml`.
3. `test_learning_session_2_2_1.py::test_compute_cwf_target_still_resolves_after_entity_cost_benchmarks` — `compute_cwf_target` returns None after the entity-cost benchmarks land.

`test_web_server.py::test_post_message_free_text_async_then_poll` fails intermittently — a `KeyError` race on `_MESSAGES` in the async worker thread, not a regression.

**Why:** documents that the post-Stage-26 nav uses `data-surface=` in dropdowns, not `data-route=`, so future nav/test work matches the real structure — and pins what "green" actually means, since several stage-era frontend suites have gone stale as surfaces were removed.

**How to apply:** more than those three = a real regression, so bisect rather than assume staleness. The 2026-08-21 finance rework retargeted `test_stage_07_finance_frontend` and `test_stage_10_frontend` at the new two-tab page and pruned 29 assertions about removed surfaces from `test_stage_11_frontend`, `test_stage_13a_overview_fixes` and `test_unmapped_categories_ux` — if those go red, check whether a surface was re-added rather than "fixing" the test. GOTCHA: `pytest.ini` sets `addopts = -ra -q`, so there is NO "N passed" summary line; count `^FAILED` lines instead. Relates to [[bug-sweep-2026-05-28]], [[finance-income-control-center]].

As of 2026-08-25 the finance suites are: five JS suites (158 checks) and four HTTP smokes — `uf_income_test` 22, `uf_expense_test` 33, `uf_setup_test` 64, `uf_import_smoke` 10. The HTTP ones drive the LIVE data.db and tear down after themselves: run them ONE AT A TIME, never chained behind a shell timeout.
