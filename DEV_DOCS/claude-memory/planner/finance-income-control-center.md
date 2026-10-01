---
name: finance-income-control-center
description: "2026-08-21 Jennifer /finance rework — two tabs, uf_income_sources is the single source of forecast income"
metadata: 
  node_type: memory
  type: project
  originSessionId: 2dcea4f4-a99e-473f-8561-1294e0a958b7
  modified: 2026-08-21T20:44:17.349Z
---

2026-08-21: Keith asked to strip the noise out of Jennifer's finance surface and
replace the hardcoded income forecast with something he drives himself. Three
decisions he made when asked, which future work should not quietly undo:

1. **Come With's own books moved to the Come With app.** The CW tab and its P&L
   grid are gone from `/finance`. The PERSONAL side stays on purpose — money he
   fronts on his personal card is real personal cash out, so "Fronted for Come
   With" and "Owed back by Come With" are still reported. CW imports (PayPal /
   Bluevine) still post to the DB silently so `uf_export_cw` / `uf_push` keep
   feeding the site; they just aren't rendered here. See [[comewith-repo-and-push-auth]].
2. **Analysis, Audit and Action items were removed — he does not use them.** Their
   `/api/finance/*` endpoints are still live for the chat surface; only the page
   is gone.
3. **Summary became a control center.** Two tabs total: *Control center* and
   *Budget*. One cash chart (net bars + reserve line, merged from the old trend +
   cumulative cards), one this-month card, the review queue, and everything else
   (reserve waterfall, savings top-ups, scenario sliders) folded into a `<details>`.

**Why the income piece matters:** forward income used to be assembled from three
places that could not be edited together — `uf_runway_inputs.ui_monthly`/
`ui_last_month`, the `budget` column of `uf_cw_earnings`, and a browser-only
"earned income" slider. He stopped reinvesting dividends and has booked gigs, and
had nowhere to put either.

**How to apply:** `uf_income_sources` (scripts/uf_income.py) is now the ONE source
of every forecast dollar — name, amount, cadence (one_time/monthly/quarterly/
annual), start/end month, confirmed-vs-expected, on/off. It is materialized into
the Personal `Income` row of `uf_budgets` from the live month through the expense
horizon, and `uf_model.runway` reads it for forecast months, so budget and runway
cannot disagree. The old UI/CW-earnings inputs were seeded into it as ordinary
lines and are no longer read; `ui_monthly`/`ui_last_month` were removed from the
runway-inputs editor deliberately. The Income row of the budget grid is read-only
(class `derived`). The month resolver exists twice on purpose — Python
`INC.amount_for` and JS `UF.incomePays` — so the panel re-totals without a round
trip; `tests/test_stage_10_frontend.py` cross-checks them and must keep doing so.
GOTCHA: `materialize_budget` is a deliberate no-op when the table is EMPTY ("not
set up yet" ≠ "zero income"); lines that exist but are all switched off do write
zero. Related: [[unified-finance-model]].

**Since superseded in shape** (still true about income itself): the page is three tabs now, not two. See [[pending-handoff-note]] and [[finance-setup-tab]].
