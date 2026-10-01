---
name: cw-cash-reserve-inflows-bug
description: "2026-09-28: CW cash reserve read low because the importers post money OUT but not IN; Bluevine credit-swallowing fixed, $100 PayPal wash + $0.79 still open"
metadata: 
  node_type: memory
  type: project
  originSessionId: 21c71d6f-845a-4732-ab58-5d2a0b40974b
  modified: 2026-09-28T16:26:39.372Z
---

**The Come With cash reserve read $2,498.12 against a real Bluevine balance of $2,956.30.** The bank reconstructs to the penny from `data/__bluevine_imports_comewith/processed/*.csv` (the statements carry a running `Balance` column — that column is the authority, not any model): opening 5,000.00 + credits 357.39 − direct debits 1,340.77 − PayPal-routed debits 1,060.32 = 2,956.30, tie-out 0.00.

**Root cause: outflows are automated, inflows are not.** `v_cash_position.revenue_in` was literally `0` — the site had never recorded a single receipt into bank or PayPal.

**FIXED 2026-09-28 — Bluevine swallowed every credit.** `uf_bluevine_ingest.py` had one `credit` branch that counted and `continue`d, justified by "the opening 5,000 is capital, not revenue". Right for the TD Bank transfer, but it silently ate a **$357.39 Stripe payout** on 2026-09-08. The rule is now **inverted**: a credit is revenue UNLESS it matches `rules/bluevine_capital_credits.txt` (currently just `transfer from td bank`). Wrong-in-this-direction is the safe direction — an unrecognised credit shows up as income you can reclassify instead of vanishing. Credits post as a positive CW row in envelope `Income`, which `uf_export_cw` now maps through `INCOME_BUCKET_TO_CATEGORY` (the expense fallback "Operations" is nonsense on a revenue row). Re-imported + pushed; reserve should read **$2,855.51**.

**PayPal is a pure pass-through — every payment out is funded by an equal "Bank Deposit to PP Account".** Bank-funded PayPal = $1,060.32 = real PayPal spend, exactly. So PayPal's own balance never funds anything, and any excess in `cash_source='paypal'` is a double-count.

**RESOLVED 2026-09-28 — the $100.** On 2026-09-08 PayPal shows `Keith Berkman Mobile Payment +100` then `Keith Berkman Payment Refund −100` (a wash). The importer posted the −100 refund as CW spend (positives go to the review queue, negatives post), and the +100 was resolved out of review to Personal / Income. **Keith confirmed the $100 was his own money and asked for it on Come With.** The row was moved `entity='Come With'` and its `source` restored from `review_resolve` to its true origin `paypal_ingest` (which is what gives it `cash_source='paypal'` on export). Both legs are now on the books, so the net float effect is zero — but note **gross** CW money-in and money-out are each $100 higher, and it books as income rather than owner capital, because `uf_transactions` has no capital mechanism for CW and `ingest-finance` only writes income/expenses/budget_lines.

**TIED OUT EXACTLY 2026-09-28 — `cash_reserve` = $2,956.30 = the bank statement.**

The $0.79 unravelled the whole thing: prod's 11th PayPal row was a hand-entered **$3.61 "Platform fees"** dated 2026-08-16, the same day as a **$361.00 Production fee** income row. `361.00 − 3.61 = 357.39` — **exactly the Stripe payout**. So the Stripe credit was *settlement of revenue already on the site's books*, not new revenue, and auto-posting it as income double-counted the P&L even though the cash figure looked right. Keith had marked that $361 `received` an hour before the push, which is the corroborating signal.

**So the Bluevine rule was inverted AGAIN, correctly this time: non-capital credits go to the REVIEW QUEUE, never auto-posted as income.** Only a human can tell settlement from new money. (A re-import of the 2026-09-23 Bluevine CSV will now enqueue that Stripe credit for review — dismiss it, it is already booked as the $361.)

Final adjustments: deleted the $357.39 income (Jennifer + site soft-delete); `361.00` income → `cash_source='bank'`; `3.61` fee `paypal`→`bank` (Stripe netted it from the payout, it never touched PayPal); the two 2026-09-05 Bandcamp $1.00 rows re-provenanced `review_resolve`→`paypal_ingest` so they carry `cash_source='paypal'`; added a **$0.82** row for the net of PayPal's four 2026-09-05 FX lines (`general currency conversion` is excluded per-line as a transfer, which drops the net cost).

**Cross-checks that all hold:** Jennifer's `paypal_ingest` total = **−1,060.32** = the bank's PayPal funding exactly · site bank spend 1,344.38 = 1,340.77 debits + 3.61 fee · site paypal spend 1,160.32 − 100.00 in = 1,060.32 · `5000 + 461.00 − 2,504.70 = 2,956.30`.

**STILL OPEN — $250.00, 2 rows, Keith's call:** `2026-08-16 Henry $150.00` and `2026-08-16 Berky $100.00`, both site-native (no `external_ref`), `status='paid'`, no cash source. Neither appears in the Bluevine or PayPal statements, so they were paid from somewhere else. They sit outside the reserve rather than being guessed at. Also `$1,200.00` income (2026-09-21) is still `invoiced`, correctly excluded.

**SUPERSEDED — $0.79.** Prod's `cash_source='paypal'` total (1,161.11) exceeds Jennifer's (1,159.50); there is an 11th prod PayPal row. Undiagnosed — `SBP_PAT` was revoked mid-session.

**Related bug found, NOT fixed:** `review_resolve` is not in `uf_export_cw.CASH_SOURCE`, so every review-queue row Keith resolves exports with `cash_source = null` and falls outside the cash reserve entirely. That is the view's `unknown_source_rows` (4 rows / $252.00); two of them are the 2026-09-05 Bandcamp charges that really were PayPal. The review queue doesn't store which feed an item came from, so fixing it properly needs provenance on `uf_review_queue`.

**`SBP_PAT` in the Comewith `.env` is revoked** (Management API 401; it worked earlier the same session). Needed to read prod. It is a Supabase **Personal Access Token** (account-level, renewed at supabase.com/dashboard/account/tokens), not a project key. Always pass `SBP_REF=yaytdosxfhcqatmhctzk` as a LITERAL — the repo's CLAUDE.md requires it, `SBP_REF=$SBP_REF_PROD` silently fails from the Bash tool, and the bare `SBP_REF` sitting in that `.env` points at STAGING (db.py's `normalize_ref()` parses the URL form fine — the value itself is the wrong project). The `PUSH_TOKEN` path (ingest-finance) is unaffected and still works.

**BUILT 2026-09-28 — settle before add** (branch `finance/settle-before-add`, commit `0130714`, NOT merged; master auto-deploys):
- **`ingest-finance` v17 deployed** (`--no-verify-jwt`). Each row walks an ordered hierarchy, first fit wins: **S0** identity → **S1** settle an open accrued/invoiced payable → **S2** same-day adopt → **S3** settle within **90 days** → **S4** insert. Ambiguity never falls through to insert — it queues. Vendor is used to CONFIRM a match, never to make one (shared token ≥3 chars; `vendor_actor_id` is what makes "Henry"/"Henry Zaradich" work).
- **S0 must never overwrite `date`** when the site row is event-linked or has `settled_at`. Jennifer sends the cash date; the site wants the incurred date. Verified live: the Henry row survived a full push still dated 2026-08-16.
- **Migration 207**: `ingest_queue` (4 reasons: ambiguous_settlement / amount_variance / vendor_mismatch / missing_ref) + `ingest_runs` (every run, incl. `report_only`). Both admin-only, anon-revoked, RLS with a real policy. A missing `external_ref` is now QUEUED, not `skipped`.
- **`report_only: true`** computes every decision and writes nothing but the run record. `S.push_cw(report_only=True)`.
- **Desktop tool**: `scripts/run_finance.py` + `scripts/Run Finance Import.bat`, with Desktop shortcuts *Run Finance Import* and *Finance Report (no changes)*. Prints SETTLED apart from ADDED on purpose.
- **`uf_push.TIMEOUT` 30 → 300.** The hierarchy makes a 243-row true-up take ~90s; the old timeout fired while the run was still SUCCEEDING server-side, which reads as a failed push.
- **Backlog remediation**: Henry $150 merged into the 2026-08-16 accrual (one candidate, both sides actor "Henry"); Berky $100 queued (three candidates, "Keith Berkman" has no actor). CW expenses −$150; reserve unchanged.
- **Dashboard**: Expenses tab shows last-run + open queue. **Not live until master is merged.** Classes there are `data-table`, not `hub-table`; `money`/`fmtDate`/`escapeHtml` are `const` arrows defined ~line 4382-4482.

Tests: edge function 13 → **23**; `uf_push_test` 19; JS suites unchanged.

Related: [[cw-push-wired-into-import]], [[comewith-repo-and-push-auth]]
