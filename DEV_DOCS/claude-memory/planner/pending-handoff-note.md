---
name: pending-handoff-note
description: Before any /finance development read NOTE_finance_control_center.md first — it is the current state of the whole surface
metadata:
  type: project
---

Before doing any development work on `/finance`, read
`C:\Users\Admin\Documents\Master\planner\NOTE_finance_control_center.md`. Written
2026-08-25 and **updated 2026-09-28**, it covers the three tabs, the
table-per-idea model, **eighteen invariants that will bite you** (twelve on the
surface, six on the push to the Come With site), how to run every suite, the
pytest 3-known-failure baseline plus the `uf_plan_test` 41/42 date-fragility, what is deliberately not built, and what is
open for Keith rather than for code.

**Why:** the surface changed shape several times in one session (expense plan →
setup tab → reserves card → notes). Guessing from the code costs an hour; the note
costs five minutes.

**How to apply:** surface it when the work is finance development. Don't recite it
for an ordinary finance question. The single most dangerous invariant: an envelope
name is a foreign key by value in six places — only rename via
`SET.rename_envelope`. See [[finance-setup-tab]], [[finance-change-notes]],
[[finance-page-reading-corrections]], [[expense-plan-month-control-center]].
