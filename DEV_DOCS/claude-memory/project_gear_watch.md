---
name: project-gear-watch
description: "Gear Watch — 3x/day scored resale scan for the DJ rig stolen 2026-08 (migration 146 + scan-gear-market); BUILT but NOT applied/deployed/pushed; Craigslist IS scannable via sapi.craigslist.org (only the RSS is dead), OfferUp/FB unautomatable; cron→edge auth = service-role bearer from vault"
metadata: 
  node_type: memory
  type: project
  originSessionId: ae01d6c5-bc1f-4ab9-9376-130207f0bafe
  modified: 2026-08-18T21:19:13.304Z
---

Keith's DJ rig (~$15K) was stolen from a vehicle overnight, 2026-08. NYPD grand larceny,
detective + evidence team. Two artifacts, both built 2026-08-18 on the desktop:

**Loss schedule** — `Financial/ComeWith_Stolen_Gear_Loss.xlsx`. Serves detective + DA
restitution + insurance. Final roster **10 units, $11,713.07** documented price paid:
XDJ-AZ ×1, CDJ-3000 ×2, CDJ-3000 case ×2, Wave 8 ×2, KRK monitor ×1, Wave 8 stand ×2.
Source: `Financial/ComeWith_Work_Expenses_Master.xlsx` → *Equipment (Capital)*.
⚠ **That sheet's D/S/C/M/W/A codes are internal ASSET TAGS, not manufacturer serials** —
police can't enter an asset tag into NCIC, so all serials still read NEEDED. Sweetwater
invoices carry serials and cover 6 of the 10 items.

**Gear Watch** — migration `146_gear_watch.sql` + edge fn `scan-gear-market` +
master-only dashboard panel + `DEV_DOCS/GEAR_WATCH.md` runbook. **HELD: not applied to
prod, not deployed, not pushed** (master auto-deploys; Keith green-lit the build, not the
deploy). Reverb + eBay have **never been called with live credentials**; Craigslist has,
and works.

Hard facts, tested 2026-08-18 — don't re-litigate:
- **Craigslist IS scannable.** Its search *RSS* is dead (403 on every `format=rss` path,
  any user-agent) and the HTML search page is a JS shell with zero listings — but the
  endpoint its own search box calls answers 200:
  `sapi.craigslist.org/web/v8/postings/search/full?batch=<areaId>-0-360-0-0&cc=US&lang=en&searchPath=sss&query=<q>`
  (NYC area id = **17**). Same trick as Bandcamp in `track-sources`. Live-verified.
  ⚠ I first concluded it was impossible; Keith corrected me and was right — when one
  access path 403s, check the site's own network tab before declaring it unscannable.
  LEARNINGS §27.
- **Its payload is delta-encoded, with three traps:** ids accumulate from
  `decode.minPostingId`; dates are seconds after `decode.minPostedDate`; `price: -1` means
  "no price stated"; the geo string is `"<locationIdx>:<descriptionIdx>~lat~lon"` — **two
  different indexes into two different arrays**; and a **zero-result query returns
  `decode: 0`** (the number), which must return `[]` and NOT read as blocked. Canonical URL
  = `craigslist.org/view/d/<slug>/<token>` from the `[6,…]` slug + `[13,…]` token.
- **Geo terms must match on word boundaries** — `"ny"` matches inside **albany**, which
  scored every upstate listing as local. `isLocalTerm()`.
- **Reverb needs TWO calls.** Its SEARCH payload has no seller location and no feedback
  (`shop` = {slug, preferred_seller} only); the DETAIL endpoint `/api/listings/<id>` has
  `location.display_location` + `shop.feedback_count`. Scan = search → model gate →
  detail lookup for survivors only (capped 40/run, cap reported). Seller name is
  `shop_name`, NOT `shop.name`. Token: reverb.com/my/api_settings → My Profile → API &
  Integrations; scopes `public` + `read_listings`; **tokens never expire**.
- **Negative signals** (added after seeing real Reverb inventory): priced >1.5× market
  −10, established dealer (preferred_seller badge) −10. Both lower a score, neither gates.
- **OfferUp + Facebook Marketplace: no public API, scraping prohibited.** Emitted as
  saved-search links, never scraped.
- **Score model** (`scoring.ts`, 33 tests, no credentials needed): gates = model named /
  not an accessory / posted after theft date; then base **35**, serial +60,
  bundle-multi-model +30, local +25, recency **graduated** (+25 ≤3d, +20 ≤7d, +15 ≤14d,
  +10 ≤30d, +5 ≤60d), pair +18, price-below-market +15/+8, new seller +10, listing
  language +10 cap. Every award stored in `score_breakdown` — a bare score is useless to a
  detective. **Recalibrated 2026-08-18** after Keith flagged a real Brooklyn CDJ-3000
  listed the day after the theft scoring only 55; it is 85 now. Thresholds: email 65,
  push 85. Model+recency without local = 60 and deliberately does NOT email.
- **Serials are the highest-value input** (+60, and the only thing that makes a recovery
  attributable).
- **The rule on a hit:** send link + breakdown to the detective. Never contact the seller,
  never arrange a meet.
- **Exercise the whole pipeline with no keys and no deploy:**
  `node scripts/gearwatch_live_test.mjs` — hits Craigslist live, prints scores, breakdowns
  and every drop with its reason. Four of this feature's bugs were found this way and none
  were visible from reading the code.

See also [[feedback-no-broad-anon-grants]], [[feedback-prod-migration-apply]].
