---
name: finance-page-reading-corrections
description: 2026-08-24; money in/out is gross not net-positive-months, income shows as UI-Work/Gigs/Passive buckets, CW capital is a receivable not reserve depletion
metadata:
  type: project
---

Three ways the /finance page reads the same ledger, all corrected 2026-08-24.
Each is a split that must tie back to the thing it splits — that invariant is
what `tests/test_finance_reading.py` exists to hold.

- **Money in / money out** are the GROSS sides of each month
  (`series[].p_in_actual` etc.), summing to the net. They used to be the sum of
  net-positive *months*, which reported $18K of income over a year that actually
  took in $93K.
- **Income shows as buckets** on the month-over-month grid — UI / Work, Gigs,
  Passive Income, Other. Still ONE `Income` envelope underneath: budgets split by
  `uf_income_sources.category`, actuals by payee (`INC.bucket_of_payee`), both
  made to tie to the envelope. Grid rows are derived off `r.section === "Income"`,
  never the row name — buckets aren't envelopes, and a name check made them
  editable.
- **Come With capital never touches the personal reserve** (final form, Keith's
  call 2026-08-25). Closed months roll forward on `personal_actual_ex_cw`; forecast
  months on `personal_expense_budget_ex_cw`; COME WITH's own balance carries that
  spend, which is why its reserve goes negative by what it owes. One reserve
  number, "Lent to Come With" as a memo under the waterfall total.

**How to apply:** don't reintroduce a name-based check for income rows, and don't
fold the receivable into the runway's $0 date without Keith deciding that
explicitly. Short payee tokens ("ui", "gig") are tokenized, not substring-matched
— `\b` in a regex has been mangled to a literal backspace twice by this repo's
tooling. See [[expense-plan-month-control-center]] and
[[finance-income-control-center]].
