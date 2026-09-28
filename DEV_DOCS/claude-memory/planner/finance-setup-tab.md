---
name: finance-setup-tab
description: 2026-08-25; /finance Setup tab (uf_setup.py) edits categories, routing rules and baselines; envelope names are foreign keys by value so renames must cascade
metadata:
  type: project
---

`/finance` has a third tab, **Setup**, backed by `scripts/uf_setup.py`: categories
(rename, Fixed↔Variable, reorder, add, delete/merge), routing rules (importer payee
+ category, and `income_bucket` rules that override
[[finance-page-reading-corrections]]'s payee bucketing), and the runway baselines
including `savings_base`, which was never editable before.

**Why it matters:** Keith wants to manage this from Jennifer, not from Claude Code.
Treat "he needs a developer for X" as a bug from here on.

**How to apply:** an envelope name is a foreign key BY VALUE across
`uf_budgets.line`, `uf_transactions.envelope`, `uf_expense_plan.envelope` and
`uf_rules.envelope`. Never patch a rename anywhere but `SET.rename_envelope` — a
partial rename orphans rows under a name nothing displays and both halves look
plausible. `set_type` must write `section` alongside `type` (budget rows carry
their own copy). `Income`, `Come With` and `Come With Fitness` are locked because
`uf_income.INCOME_LINE` / `uf_model.CW_ENVELOPE` find them by exact string; the
lock ships its reason so the UI can explain rather than just refuse.

**Savings**: `savings_base` in Setup is the OPENING balance for the runway window
— editing it restates closed months. A dated change is a `uf_reserve_topups` row
with `kind='savings_adjust'` (signed, reserve untouched), recorded from the reserve
card. Don't conflate it with `kind='to_reserve'`, which moves savings into the pot.

**Reserves & capital** is one card on the control center (`#pots_card`, payload
`data.pots`) covering all four pots — personal reserve, savings, Come With reserve,
lent to Come With — each opening → what moved → now, plus a dated movements ledger.
Four movement kinds in `uf_reserve_topups.kind`: `to_reserve` (savings→reserve,
does NOT reduce what CW owes), `savings_adjust`, `reserve_adjust` (moves the cap
with it), `cw_repayment` (the business paying back — reserve up, invested down).
Come With's own reserve is not a tracked pot: off the card, off the sliders, out of
the runway headline. `Invested in Come With` is ONE figure = opening float +
everything since − paid back. The reserve waterfall shows money in AND money out,
never a netted "living" line. Come With capital is
derived from charges filed to the `Come With` envelope, never typed; its window must
match `runway.start` or the running total contradicts the headline.

**Don't bury things in the `<details>` accordion.** Keith went looking for a card
there and couldn't find it. Serving something is not showing it.

The plan horizon **rolls** now (`EXP.plan_horizon()` = later of 2027-12 and 18
months out) — don't reintroduce a fixed end date, and don't hardcode runway row
counts in audits. See [[expense-plan-month-control-center]].
