---
name: project_pricing_tool
description: "Sales \"Pricing\" quote tool (DJ/rental/full-production) — DEPLOYED (migrations 068+070 applied, pushed). Quotes link to events; travel = mileage + drive time"
metadata: 
  node_type: memory
  type: project
  originSessionId: 59fd8c20-e65f-4382-bad4-d875b13df3be
---

Dynamic event pricing/quote tool in the dashboard Sales group, between Inquiries
and Agreements (built + DEPLOYED 2026-06-28). Migrations **068** (pricing_config) +
**070** (events.quote jsonb) applied to prod; frontend pushed; live on Netlify.

- **Link to event:** dropdown of events; "Save to event" writes the quote (inputs +
  totals) to `events.quote` jsonb; the linked event shows on the quote/print/copy and
  a saved-quote indicator appears when re-selected.
- **Travel** (replaced flat delivery): `cfg.travel` = mileage (round-trip beyond a free
  radius, ~$0.70/mi) + drive time (round-trip hours @ $50/h) + optional base; uses the
  Event section's Distance + Drive-time inputs. Setup/strike kept.
- **Deposit %=0** now removes the deposit line (was falling back to the default via `||`).

- **Engine:** `assets/pricing-engine.js` (pure ES module — `PRICING_DEFAULTS`,
  `computeQuote`, `mergeConfig`, `suggestDailyRate`). Imported by dashboard.html AND
  `scripts/test_pricing.mjs` (13 scenario tests, all pass — `node scripts/test_pricing.mjs`).
- **Defaults** (small DJ/production co., 2025-26 market): DJ tiers $400/$750/$1500,
  hourly $150; labor tech $65/hr, day $600, half $350, OT 1.5×; delivery $75+$0.50/mi
  past 15mi free; lighting $500/$999/$2000; rental = inventory daily_rate (unpriced =
  10% of cost), extra days 50%, 8% waiver; weekend+15/peak+20/rush+15/deposit 50.
- **Equipment** rates pulled live from `equipment_inventory.daily_rate` (only 9 of 79
  priced; rest get the 10%-of-purchase suggestion, flagged "est").
- **Storage:** migration `068_pricing_tool.sql` = `pricing_config` (single-row id=1,
  jsonb, admin-only RLS) + a `module_registry` row key='pricing' sort 15. App deep-merges
  saved config over engine defaults; per-DJ overrides in `config.dj_overrides{actorId:rate}`.
- **UI** (`loadPricing`/`renderPricing`/`computeQuote`): collapsible builder (Event / DJ /
  Rental / Production+labor / Adjustments / ⚙Edit defaults) + sticky live quote, Copy
  summary + Print/PDF. Every prefilled rate editable per-quote; defaults editable+saved.
- **renderNav has a client-side fallback** that injects the 'pricing' module for master
  if the DB row is absent (so it previews before 068). canSave=false (preview) when the
  pricing_config table is missing.
- **To deploy:** apply 068 (Mgmt API w/ SBP_PAT, see [[feedback_prod_migration_apply]]),
  push dashboard.html + assets/pricing-engine.js. See [[project_email_campaigns]] for the
  campaign edit+CC work shipped same day.
