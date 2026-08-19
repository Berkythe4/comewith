# Memory index

- [Unified finance model (uf_*)](unified-finance-model.md) — Jennifer DB canonical for both budgets; uf_* tables in data.db, Excel mirror, CW→Personal income link
- [Personal.xlsx Simplifi importer](personal-xlsx-simplifi-import.md) — Excel-only pipeline (scripts/import_personal_simplifi.py), firewalled from Jennifer's DB
- [uf_ingest Work Expenses routing fix](uf-ingest-work-expenses-routing-fix.md) — 2026-07-07; Simplifi Work Expenses + CW watchlist now post to Come With envelope (fixed June ~$0 → −$3,933), backfilled Apr–Jul
- [PayPal + manual entry](uf-paypal-and-manual-entry.md) — 2026-08-18; CW business PayPal importer (CW-only, no personal mirror), vendor learning, hand-entered charges, one import button for both inboxes
- [Come With repo + push auth](comewith-repo-and-push-auth.md) — CW website repo at Documents/Comewith; master AUTO-DEPLOYS; ingest-finance token auth shipped on a branch 2026-08-18
- [Bug sweep 2026-05-28](bug-sweep-2026-05-28.md) — all 5 open jennifer_bugs fixed (#548/#549/#550/#551/#563)
- [Pre-existing frontend test failures](preexisting-frontend-test-failures.md) — Stage 26 nav uses `data-surface=` dropdowns not `data-route=`; 8 stale frontend tests fixed 2026-05-28, only the date-relative refiner flake remains
- [CWF research R6–R12 continuation](cwf-r6-r12-continuation.md) — done 2026-06-10; key finding: $330k CAPEX understated, real ~$700k–$1.1M for the chosen warehouse footprint
- [CWF deep-verify revisions (DV1–DV9)](cwf-deepverify-revisions.md) — done 2026-06-10; 6 revisions to R1–R12: build cost ($290k=2nd-gen floor), timeline (6–12mo not 2), industrial vacancy (~3% tight not 21%), retention (~65% base), utilities (~$30k), cash-to-breakeven ($1.1–2.0M)
- [Come With ownership & sweat equity](comewith-ownership-and-equity.md) — sole owner today; Martin/Henry/Janelle each 5% over 2 years, none vested
- [Come With 1099 tracking](comewith-1099-tracking.md) — per-payee not per-category; reportability is a stored decision, not inferred
