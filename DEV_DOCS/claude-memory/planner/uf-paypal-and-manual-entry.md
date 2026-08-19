---
name: uf-paypal-and-manual-entry
description: 2026-08-18 — PayPal (CW business account) import source + hand-entry of charges on the Jennifer finance page
metadata: 
  node_type: memory
  type: project
  originSessionId: 8e5a2c45-91f9-4945-b2e7-2e39db964e9f
  modified: 2026-08-18T18:49:11.508Z
---

2026-08-18: added a second import source and manual charge entry to the unified finance model.

**PayPal = Come With BUSINESS account** (Keith confirmed: own balance/bank, NOT funded by his personal card, not in Simplifi). So `scripts/uf_paypal_ingest.py` posts to the **Come With entity ONLY — no Personal mirror**. This is the key difference from `uf_ingest.py`, where a business charge on the personal card is both a personal cash outflow and a CW expense. Inbox `data/paypal_imports/` → `processed/`. Outflows → `uf_transactions` (Indirect engine, `CW_MIRROR_NOTE`, bucket via `cw_bucket()`); **inflows → review queue** (Keith says outgoing-only today but expects income later, and revenue can't be auto-assigned to a gig/party line); non-USD → review; bank↔PayPal transfers skipped via `rules/paypal_type_exclusions.txt`; non-Completed skipped.

**Why:** Keith pays CW vendors from PayPal, so that spend was invisible to the CW P&L.

**How to apply:**
- Dedup key is PayPal's **Transaction ID** (stronger than the Simplifi occurrence-hash; that's the fallback when absent). Extra guard: a PayPal row matching an existing `cw_actuals_ledger`/`cw_ledger_archive` row on (date, amount) is skipped as `ledger_dup`.
- Handles BOTH PayPal layouts — current (`Amount`) and legacy (`Gross`/`Fee`/`Net`); **Net beats Gross** so fees are reflected.
- **Vendor learning:** `cw_bucket()`'s keywords were built for card retailers, so PayPal vendors (rental houses, venues, freelancers) land in `Other` first. Re-bucketing a `paypal_ingest` row in the drill-down calls `uf_paypal_ingest.remember_vendor()`, which writes it to `rules/paypal_vendor_map.yml` (preserves comments, updates in place) so later imports self-classify. Deliberately does NOT guess.
- **One button for both inboxes:** `uf_server.run_all_imports()` drains Simplifi then PayPal; `/api/uf/import` now returns rows tagged with `source`.
- **Manual entry:** `uf_server.add_manual_txn()` / `delete_manual_txn()`, sources `manual` / `manual_cw_mirror`. Dual-posts when entity=Personal + envelope='Come With', or entity='Come With' + `mirror_personal` (the "I paid out of pocket" checkbox). Warns (doesn't block) when a same-shape charge already exists. Delete is restricted to manual sources and removes the dual-posted pair.
- New routes in BOTH `scripts/uf_server.py` and `src/web_server.py`: `GET /api/uf/buckets`, `POST /api/uf/txn/add`, `POST /api/uf/txn/delete`. UI: `+ Add expense` button + `#uf-add-modal` in `static/finance.html`, logic in `static/uf_dash_module.js`, styles in `static/finance.css`.
- The `Run import` button has a hover tooltip documenting both inbox paths.

**Round 2 (same day):**
- Inboxes RENAMED by Keith to sort to the top: `data/__simplifi_imports` and `data/__paypal_imports` (double underscore). Constants updated in `uf_ingest.py` / `uf_paypal_ingest.py`; tooltip + READMEs follow.
- **Import summary is now a persistent on-page panel** (`#uf_implog`, `.uf-implog` good/warn/bad/busy states) instead of a browser alert — per file: added / to-review / dups / transfers ignored, plus a headline and a Review-queue pointer. Dismissible. Manual adds report into the same panel saying which books the charge hit.
- **Refresh button** (`#uf_refresh` -> `refreshData()`, re-GETs `/api/uf/data`).
- **P&Ls already auto-update** — `uf_dashboard_build` computes every actual with live `SELECT ... FROM uf_transactions` and NO source filter, so manual/PayPal rows flow in with no extra wiring; `add_manual_txn` returns fresh `_data()` and the JS re-renders in place.
- **`months_all` extended 21 -> 33 months** (2025-04..2027-12, matching `series`) in `uf_dashboard_build.build_data`. It previously stopped at 2026-12, so a charge logged into 2027 would sit in uf_transactions but be INVISIBLE in both grids.
- **By-bucket popup:** `txns()` takes an optional `section` arg (Personal grid section via `_TYPE_SECTION`, or a CW engine) returning every charge under it with `grouped:true`; the frontend renders banded per-bucket blocks with subtotals, biggest spend first. Section actual cells are now drillable (`data-section`), env cells unchanged. Verified popup totals tie out EXACTLY to the grid cell for 2026-06/07 on both entities.
- `scripts/uf_module_test.js`: its `makeEl()` DOM stub lacked `querySelector`/`classList.add|remove` and broke on the new panel code — extended the stub (not the production code); now 26/26.

**Round 5 (2026-08-18) — Option A (Personal-only), reserve gap breakdown, CW handoff decided:**
- **ARCHITECTURE DECIDED: Come With accounting moves to the Come With website.** Jennifer = bank connector + Personal only. Flow is **Jennifer → site** (outbound push of CW-shaped transactions), NOT site → Jennifer. True-up is a ONE-TIME historical migration of 180 CW rows (−$14,554.13: 92 `simplifi_cw_mirror`, 52 `cw_actuals_ledger`, 30 `cw_ledger_archive`, 6 `paypal_ingest`) + 25 gig/forecast budget lines + 6 `uf_cw_earnings`. Plan artifact: https://claude.ai/code/artifact/af973e81-7e33-4f36-b957-ca1ae3467bd1
- **Keith chose Option A**: Jennifer drops the combined runway and the CW reserve from its headline. KPI row is now Personal reserve / last month's net / owed back by Come With / reserve runs out. `combined_zero` is still computed but is no longer surfaced. Chart default flipped `combined` → `personal`.
- **What can NEVER leave Jennifer**: the 182 Personal rows tagged `Come With` (−$18,996.62). They are personal cash out AND invested capital, and they drive the reserve/runway.
- **Reserve gap decomposition** (the round's main ask — "wherever we show personal reserve also show how much of the gap is invested"): `uf_model.runway` rows now carry running `gap_living / gap_invested / gap_interest / gap_topup / gap_capped`; `kpis.gap` exposes them. Reconciles exactly: `opening + living − invested + interest + topup − capped == reserve_p`. Verified 50,000 − 5,276.68 − 6,027.07 + 841.70 − 8,236.28 = 31,301.67. Rendered as a waterfall (`#reserve_gap`, `.uf-waterfall`) plus a sub-line on every reserve card via `reserveSub()`. "Re-up" renamed **Replenish**.
- **GOTCHA — two invested figures, both correct**: `gap.invested` is as-of `actual_through` (matches the reserve), `outstanding_cw` is through TODAY (money already spent isn't a projection). Both now carry explicit as-of labels; do not "fix" one to match the other.
- **`PENDING` reverted** — receivables belong to the website under this architecture. `uf_transactions.status` holds only NULL and 'REVIEW' (274 migrated personal rows).
- Keith is running the Come With website work in a SEPARATE workflow/chat.
- Test gotcha: suites that read the live review queue break once it's emptied — seed your own rows instead (the queue legitimately hit 0 after he filed all 4).

**Round 4 (2026-08-18) — reserves as capped pots, top-ups, actionable review queue:**
- **A reserve is a defined POT, not a bank balance.** Each is CAPPED at its opening balance (Personal 50k, Come With 5k) in `uf_model.runway`. Anything above the cap spills into `savings` instead. Keith's call: "cap at 50", "the come with reserve should never be over $5K".
- **Come With float**: only `cw_outside_spend()` (= `cw_actual_net + cw_investment`, i.e. PayPal / CW ledger / CW-only manual) draws it down. Personal-card CW cost arrives as capital and is spent the same month → nets to zero. **No interest on the CW float.** Verified: 5000 → 5000 (May, all personal card) → 4800 (Jun, −200 PayPal) → 4500 (Jul, −300 PayPal).
- **Interest**: personal only, and the base is `(reserve_p + savings)` — but only the INTEREST lands in the reserve, never the savings principal. New input `savings_base` (50000).
- **Top-ups**: new table `uf_reserve_topups(month, entity, amount, note, created_at)`. `uf_server.add_topup/list_topups/delete_topup`, routes `GET /api/uf/topups`, `POST /api/uf/topup/add|delete`. A top-up moves money from savings into a reserve (still capped) and reduces `outstanding_cw`. UI: "Reserves and capital" card with a re-up control + undo list.
- **invested/outstanding count through TODAY**, not through `actual_through` — unlike the reserve these aren't projections, the money is already spent. `invested_cw = cap_cw + invest_to_date`; `outstanding = invest_to_date − topups`.
- **Review queue reworked from read-only to actionable**: `uf_server.review_queue()` returns items + per-entity bucket choices + a parsed `category`; each row in the UI gets books/envelope selects, a "remember" checkbox (writes a `uf_rules` payee rule), **File it**, and **Ignore** (`dismiss_review` → status `IGNORED`, kept not deleted). `apply_review` now also returns the refreshed queue.
- Zero dates shifted earlier once the May surplus stopped inflating the reserve: personal $0 ≈ 2027-01, combined ≈ 2027-02. NOTE the reserve can go negative while `savings` is still large — the zero date means "reserve exhausted", not "broke"; savings is surfaced as its own KPI so that isn't misread.
- Keith noted the big May inflow was his final paycheck and "should count as income for April and May" — NOT re-dated; flagged only.
- Test gotcha: suites that hard-code counts by `source` break once real data of that source lands (the PayPal idempotency check). Scope such assertions to the fixture's own rows.

**Round 3 (2026-08-18) — Come With as invested capital + Summary rebuild:**
- **BUG, real cause of "nothing updates":** `uf_dashboard_build.CURRENT` was the literal `"2026-06"`. It froze the KPI cards, the this-month chart and `months_actual` on June while the ledger moved on. Now `datetime.now().strftime("%Y-%m")`. Also `data/uf_dashboard/data.json` (the JS test fixture, written by `_regen_side_effects`) was stale — regenerate it after payload changes.
- **CW cost on the personal card = capital Keith invested**, per his instruction. `uf_model.cw_investment(m)` = -(Personal rows with envelope='Come With'). It leaves the personal reserve and arrives in CW, which spends it immediately, so CW cash moves only by interest and combined cash falls exactly once. `cw_invested` accumulates (seed `reserve_cw` 5000 + contributions; currently 10,282.42).
- **Personal tab no longer shows Come With at all** — `psec` is now `['Income','Fixed','Variable']` and `pmemVals` filtered to match. The rows still exist in `uf_transactions` (real cash out, still in the runway); they are reported as investment instead.
- **Reserve rolls forward from actuals.** `uf_model.runway(..., actual_through=)` defaults to `last_closed_month()`; months ≤ that use `personal_actual_net` / `cw_actual_net`, later months forecast from budget. New payload keys: `reserve{month:{p,cw,combined,invested,actual}}`, `cw_invest`, `invest_to_date`; kpis gain `actual_through/reserve_*/invested_cw/burn_p`.
- **Summary rebuilt as a numbered story** (`.uf-story-step`): 1 where you stand (KPI cards) → range bar → 2 what moved → 3 where the balance heads → 4 this month vs plan → 5 what you spent on → 6 how long it lasts. Each chart carries its own `.kpis-mini` strip. One date range (`ST.sum`, `#sum_fromD/#sum_toD` + 6m/12m/24m/all presets) drives every chart via `UF.viewData()`.
- **Reserve line on both charts**: `netBarsSVG` draws it on a right-hand secondary scale (solid measured / dashed forecast); `cumSVG` now uses the server `reserve` instead of re-forecasting client-side. Charts split on `UF.closedThrough(D)` (last CLOSED month), NOT on today — otherwise the part-month current bar reads as "spent nothing". `UF.runway()` (scenario sliders) starts the month after `actual_through` from the MEASURED balances so it agrees with the cumulative chart.
- **Popup shows which pocket paid**: `uf_server.paid_from(source)` → personal / business / manual chip.
- **GOTCHA — two copies of the logic file.** `static/uf_dash_logic.js` (what /finance loads) and `scripts/uf_dash_logic.js` (what the standalone `uf_server.serve_html` inlines) have DRIFTED (17.5KB vs 12.5KB). `uf_module_test.js` read static/ but `require`d scripts/, so it was validating the wrong copy and hid a `UF.viewData is not a function` failure. Test now requires static/. The scripts/ copy is still stale — the standalone dashboard lacks these features.

**ANSWER to "does this update Come With's website?" — NO.** There is no website/CMS integration anywhere in the planner: the only outbound credentials in `.env` are the Anthropic API key and Google Calendar, and nothing in `scripts/uf_*` or `src/` posts to an external host. All of this is local — `data.db` plus the Excel mirror on this machine.

**GOTCHA (environment, cost me two failed edits):** the Bash tool mangles backslashes inside heredocs even with a quoted delimiter (`<<'PY'`) — `\\n` collapses to a real newline and `\U` triggers a Python unicodeescape error. Use the Write/Edit tools for any content containing backslashes (Windows paths, JS escapes).

Testing: always against a COPY of `data.db` in the scratchpad, never live. Still true — **never run `scripts/uf_server_test.py`** against live data.db (see [[uf-ingest-work-expenses-routing-fix]]). Related: [[unified-finance-model]], [[personal-xlsx-simplifi-import]].
