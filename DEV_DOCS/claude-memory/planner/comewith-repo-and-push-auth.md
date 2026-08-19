---
name: comewith-repo-and-push-auth
description: "Come With website repo location + binding conventions, and the Jennifer→site push token auth shipped 2026-08-18"
metadata: 
  node_type: memory
  type: project
  originSessionId: 8e5a2c45-91f9-4945-b2e7-2e39db964e9f
  modified: 2026-08-18T21:58:01.161Z
---

**The Come With website repo is `C:\Users\Admin\Documents\Comewith`** (GitHub `Berkythe4/comewith`, site `comewith.org`). Separate from the planner. Stack: **Supabase** (Postgres + Deno/TS Edge Functions, prod ref `yaytdosxfhcqatmhctzk`) + **Netlify** hosting + **Resend** email. Do not confuse with `Documents/Master/come-with-fitness` (CWF, the fitness venture).

**BINDING conventions there (`CLAUDE.md` — read `CARRYOVER.md` first, always `git fetch` before writing any doc):**
- ⚠ **`master` AUTO-DEPLOYS to Netlify.** Pushing master publishes the site immediately. Work not green-lit goes on a branch, named in CARRYOVER under "Parked / next".
- Edge functions do NOT deploy on push — only via `python scripts/deploy_edge_function.py <slug>` (the Supabase CLI is linked to STAGING and rejects the `sbp_v0_` token format).
- Migrations: introspect prod first, DRY-RUN every migration (`commit;`→`rollback;`) before applying. Never `grant ... to anon`. Financial views are anon-revoked by design (must return 401).
- Session close = `MERGE_ROUTINE.md`. Three machines ship into this repo (Keith desktop + laptop, Henry), Claude memory doesn't sync, so repo docs are the handoff.
- Edge-function tests run under **`node --test`** with native TS (not `deno test`) — see `scan-gear-market/scoring.test.ts`. GOTCHA: `node --test <dir>` fails on Windows here (resolves the dir as a module); pass the **test file path** directly. That affects their existing test too — pre-existing, not something I broke.

**Shipped 2026-08-18 on branch `security/push-token-auth`** (pushed, NOT merged; PR at github.com/Berkythe4/comewith/pull/new/security/push-token-auth), implementing `HANDOFF-push-token-security.md`:
- `supabase/functions/ingest-finance/index.ts` — receives Jennifer's fee/vendor push. Static bearer token, constant-time compare on SHA-256 digests, bare 401, **fails closed** if `PUSH_TOKEN` unset (500, never "auth off"). **Rejects a token in the query string** — deliberately unlike `ingest-email`, which uses `?key=` and must not be copied. 7 tests pass.
- Storage intentionally NOT wired — payload contract unagreed, and a table would need a migration that hasn't been dry-run. Endpoint accepts + counts so the token path is verifiable first.
- `.gitignore` `.env.*` + `!.env.example`; `.env.example` gains `PUSH_TOKEN` (key only) and `PUSH_ENDPOINT`; `.githooks/pre-commit` blocks the `push_[0-9a-f]{64}` pattern and any staged `.env`; `docs/ROTATE_PUSH_TOKEN.md`.
- Verified `.env` was **never** committed (`git log --all --full-history -- .env` empty) — that was the handoff's stop condition.

**Planner side:** `scripts/uf_push.py` — reads `PUSH_TOKEN`/`PUSH_ENDPOINT` from the planner's `.env`, fails fast naming the VARIABLE never the value, refuses non-https, sends `Authorization: Bearer`, whitelists payload fields (never `dict(row)`), logs status+body but **never** headers. `describe_config()` reports token length only.

**THE CRITICAL FINDING (2026-08-18) — the two ledgers OVERLAP.** Measured, not assumed:
- Jennifer: **180** Come With rows, **$14,554.13**
- Comewith prod `expenses`: **133** rows, **$26,976.42** (112 are `event_na` overhead), span 2024-12-23 → 2026-07-11
- Matched on (date, amount): **66 are the SAME charge in both** ($4,874.34) · **114 Jennifer-only** ($9,679.79, mostly Jul/Aug where the site has nothing) · **67 site-only**
- A straight import creates 66 duplicates. NEVER push Jennifer→site without the adopt logic.
- Note the site's total is ~2× Jennifer's: Jennifer's CW P&L has been UNDERSTATING real Come With spend.

**Migration 147 (`147_fpa_pl.sql`) — WRITTEN + DRY-RUN CLEAN, NOT APPLIED.** Adds `expenses.external_ref` + `income.external_ref` (partial unique), `expenses.funded_by` ('business'|'owner'), `budget_lines.period` + scope 'period', views `v_pl_monthly` / `v_pl_monthly_vs_budget` / `v_owner_funded` (all anon-revoked), and a `module_registry` row for the P&L tab. Dry-run = `sed 's/^commit;$/rollback;/'` then `SBP_REF=yaytdosxfhcqatmhctzk python db.py <file>` — verified prod untouched afterwards.

**`ingest-finance` now stores, with three-way resolution:** own external_ref → update · same date+amount with NULL ref → **ADOPT** (claim it, preserve the site's hand-entered category/vendor/event_id, take funded_by from Jennifer) · else insert. Idempotent. 12 tests pass via `node --test <file path>` — the Supabase client is behind a `__setDbFactory` seam and lazy-imported AFTER auth, so tests never hit the `npm:` specifier.

**Jennifer side: `scripts/uf_export_cw.py`** — emits all 180 rows with `external_ref = uf_transactions.hash` (already unique) + 37 budget lines. Flips sign (Jennifer stores costs negative, site stores positive). Maps buckets → the site's real category vocabulary. 92 rows owner-funded ($7,766.57 = what CW owes Keith back). `--summary` to inspect without writing.

**Dashboard P&L tab** added to `dashboard.html` (panel-pl / plBody / `loadPL()`, dispatch at the `tab === 'pl'` line), following the Gear Watch pattern including the "not installed yet — apply 147" state.

**CORRECTION Keith made:** only HIS machine ever holds the push token (planner `.env` + Supabase secrets). Henry's/Martin's clones have nothing to leak — my earlier "install the hook on all three machines" guidance was wrong and is fixed in the hook comment and runbook.

**Windows gotcha:** `push_$(openssl rand -hex 32)` is bash and fails in cmd.exe. Give him `python -c "import secrets; print('push_' + secrets.token_hex(32))"`.

**Rounds 6-9 (2026-08-19) — the P&L became real. Migrations 147-160 all applied.**
- **Bluevine IS the $5,000 float.** `data/__bluevine_imports_comewith/` + `scripts/uf_bluevine_ingest.py` (`--reconcile` = read-only report). Export opens `Transfer from TD Bank 5,000.00` 2026-06-25. THREE row kinds: credits = capital (never revenue); **PayPal-routed debits SKIPPED** (the PayPal feed already holds them, with the real payee — Bluevine only says SOCHETJANEL); direct card debits post as CW cost with `cash_source='bank'`. Found **$642.94 of spend in no ledger at all**. Settlement lag 1-3 days, so some spend is legitimately in flight (books $3,551.06 vs bank $3,651.06 = $100 Henry Zaradich not yet settled). Software switch personal→business card was CLEAN, no double-billing; **Splice is the one still on the personal card**.
- **Ledgers split**: `expenses.ledger`/`income.ledger` ('come_with'|'dance_infusion'), derived from event series. DI was **96% of "revenue"** ($12,497.94 of $13,049.83). CW's own revenue is ~$550.
- **v_pl_monthly was missing nearly all revenue** — it read only `income` (1 row, $1.89) and ignored ticketing/sponsorships/donations. Migration 022 already had the canonical basis. Sectioned now: Revenue − Direct = Gross − Indirect = Net.
- **`sum(...) filter` returns NULL not 0** — months with cost and no revenue rendered gross/net as `0.00`, the most misleading possible value. Always coalesce.
- **vendor_aliases + resolve_vendor_actor()** (longest pattern wins): 72 payee strings → 47 aliases, linked 132→256. Merging in the P&L payee view writes the alias too. Found+merged duplicate actors incl. a mojibake `Crossroads CafÃ©`.
- **Only DI has revenue recorded.** 5 past CW events spent with $0 income; `v_event_money.missing_revenue` flags them (excludes future events) and the Income tab banners it with one-click add.
- Income now has full Expenses parity; `revenue_streams` names the vocabulary.
- **P&L presets**: This month / Last 3 + now (default) / Forecast / YTD / Full year. DI tab checks itself against `v_kpi_dance_infusion` (published on the public site) — ties at $12,499.83 raised / $4,142.50 donated.

**GOTCHAS worth keeping:**
- `create or replace view` can only APPEND columns — never rename or reorder. Renames need drop-then-create (safe only for leaf views).
- `ON CONFLICT` cannot use a PARTIAL unique index unless the predicate is restated.
- Merging actors: `actor_roles` has unique `(actor_id, role)`; a blind repoint collides. Repoint every FK generically, fall back to DELETE on unique_violation.
- `fmtDate` formats date-only strings by SLICING — `new Date('2026-08-19')` is UTC midnight and local getters render the 18th in Eastern.
- **The other session churns migration numbers fast** — 150, 152, 155 were all taken mid-build. Always re-check across ALL refs immediately before applying.

**Open / needs Keith:**
1. Generate the token himself: `echo "push_$(openssl rand -hex 32)"` — never run it in an assistant session. Set on server via `supabase secrets set PUSH_TOKEN=…` AND in the planner's `.env`.
2. Deploy with JWT verification OFF: `python scripts/deploy_edge_function.py ingest-finance` (`--no-verify-jwt`).
3. `git config core.hooksPath .githooks` must be run **per clone** — laptop and Henry's machine too, or the hook is absent exactly where a leak would come from.
4. Their convention wants the branch named in CARRYOVER "Parked / next" — I did NOT edit CARRYOVER (append-heavy shared file, Gear Watch had just rewritten it, and conflicts there are the thing their rules warn about).

**Also note:** local `master` there is 1 commit ahead of origin (`83e7f5f` Gear Watch) that Keith deliberately left unpushed *because* master auto-deploys. My branch has it as an ancestor, so it's now visible on the remote as part of that branch — but master itself is untouched and nothing deployed.

Related: [[uf-paypal-and-manual-entry]] (the Jennifer→site handoff architecture, Option A).
