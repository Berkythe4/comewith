---
name: finance-change-notes
description: 2026-08-25; uf_notes records WHY something changed, hung off a category and/or month; never re-materializes the budget
metadata:
  type: project
---

`scripts/uf_notes.py` / table `uf_notes` records **why something changed** — the
one kind of note the model lacked. A note hangs off a category, a month, or both;
attached to neither it is refused. Surfaced as "Why this month looks like this"
under the month plan, plus a dot in the budget grid's month header.

**Why:** every other note describes a thing (what a line is, what a charge was
for). Six months on, "groceries went up in November because X" is the only thing
that makes a month readable.

**How to apply:** two rules that are easy to break.
1. Saving a note must NEVER call `_materialize_all` — a note explains numbers, it
   does not move them. `tests/test_uf_notes.py` asserts the absence.
2. `uf_notes.envelope` is the **sixth** place an envelope name lives by value.
   The rename cascade in `SET.rename_envelope` and the merge path must carry it,
   and `_counts` includes notes so deleting a category reports them.

See [[finance-setup-tab]] for the other five places and the cascade itself.

**Terminology is pinned by test.** `tests/test_uf_notes.py` fails if "Owed back",
"Lent to Come With", "Come With reserve", "re-ups", "Cash runs out" or "Personal
living" appear in finance.html or uf_dash_module.js outside a comment, or in the
stored `page_metadata` help. Rename in one place and the test finds the rest.

**An import must reload the side lists.** `data` carries everything derived, but the
review queue, setup lists and movements are cached separately —
`reloadSideLists()`. `scripts/uf_import_smoke.py` proves it end to end against the
live server and cleans up after itself.
