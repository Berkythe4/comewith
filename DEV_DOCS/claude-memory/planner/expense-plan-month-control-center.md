---
name: expense-plan-month-control-center
description: 2026-08-24; uf_expense_plan is the single source of forward expenses, edited on the Budget tab; amount changes ask "just this month" vs "from here on"
metadata:
  type: project
---

The Budget tab (`/finance`) carries a month plan: `uf_expense_plan` +
`uf_expense_overrides` in data.db, the expense twin of [[unified-finance-model]]'s
`uf_income_sources`. Lines sum per envelope per month and materialize into the
Personal expense rows of `uf_budgets` (negative), live month forward only, through
`EXP.PLAN_HORIZON` = 2027-12. Income is edited from the same panel, reusing the
income endpoints — not a second list.

**Why:** expenses used to exist only as numbers typed into budget cells, with no
record of why a month cost what it cost and no way to say "this changes in
December". The forecast past 2026-12 was December's budget held flat.

**How to apply:** changing a recurring line's amount has two honest readings, so
the row asks inline — "Just Dec '26" writes an override; "Every month from Dec '26"
SPLITS the line (ends it in November, starts a new one in December) rather than
editing the amount in place, which would restate months already reported on. Never
"simplify" that into a plain amount update. A planned envelope's budget cell is
shown, not typed, from the live month on; an envelope with no plan line keeps its
budget and stays editable. Seeded from Keith's Simplifi widget at $7,044.94/mo —
envelopes match `rules/personal_envelope_map.yml` (where the importer files the
charge), not where the name reads best. The widget only shows what is due THIS
month, so two things had to be added that it could not show: the twice-yearly
$1,100 insurance premium, and the COBRA step (Health $300 through 2026-10, $900
from 2026-11, read from `uf_runway_inputs.cobra_may_oct`/`cobra_nov_plus`). Don't
flatten either back to one line — that quietly makes the runway cheaper. See [[finance-income-control-center]] and
[[preexisting-frontend-test-failures]].
