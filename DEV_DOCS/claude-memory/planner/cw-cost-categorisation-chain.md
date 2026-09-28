---
name: cw-cost-categorisation-chain
description: "How a Come With charge gets its P&L category, and why named freelancers silently land in Operations"
metadata: 
  node_type: memory
  type: project
  originSessionId: 21c71d6f-845a-4732-ab58-5d2a0b40974b
  modified: 2026-09-28T18:21:50.509Z
---

**A Come With cost gets its site category through four hops, and a miss at any one lands it in Operations without complaint:**

1. `rules/paypal_vendor_map.yml` — vendor substring → Jennifer bucket. **This file ships entirely commented out**, so until a vendor is taught it matches nothing.
2. `uf_ingest.cw_bucket()` — a keyword classifier built for card retailers (Ableton, Beatport, Google Ads). **It cannot recognise a person's name.**
3. Fallback → Jennifer envelope **`Other`**.
4. `uf_export_cw.BUCKET_TO_CATEGORY` maps `Other` → **`Operations`** on the site.

So every freelancer paid by name — marketing, design, a DJ — silently reads as Operations in the P&L. **Found 2026-09-28:** Janelle Sochet's monthly marketing retainer ($200 Jun / $300 Jul / $200 Aug = $700) sat in Operations, leaving Marketing showing only a $500 Facebook charge. Fixed by `uf_paypal_ingest.remember_vendor('janelle sochet', 'Marketing')` (which writes the YAML, preserving comments) plus recategorising the three existing rows; Marketing went $700 → $1,400.

**The lesson: an empty Marketing or Contractors line is more likely a routing miss than an absence of spend.** Check `Other`/`Operations` before believing a category is empty. Valid Jennifer buckets are only: `Software | Marketing | Venue | Gear / Production | Other` — the site's vocabulary is richer (Contractors, Platform fees, Travel, Professional Development…), so the mapping is lossy by design and anything unmapped collapses to Operations.

`remember_vendor` is also called automatically when a PayPal row's envelope is changed in Jennifer's drill-down (`uf_server.update_txn`), so reassigning once teaches the importer.

**THE 66-DUPLICATE PROBLEM WAS NEVER ACTUALLY SOLVED (found 2026-09-28).** `ingest-finance`'s adopt matched on **exact date + exact amount**, so it caught only the cleanest collisions between the site's hand-kept ledger (57 rows created 2026-05-29) and Jennifer's 2026-08-19 push. Everything with settlement lag or an FX difference duplicated silently.

Measured: of 62 unreffed CW expenses, **26 pair to a feed row within 0–5 days at a median amount ratio of exactly 1.0664** — a constant ~6.64%, i.e. the same charge booked once at the quoted (EUR) price and once at what the card was actually billed. **$905.71 of double-counted cost**: Beatport 23 pairs/$637.75, Ableton 2/$67.96, Janelle 1/$200.00. Three pairs are ratio 1.0000 (identical amounts, 1–2 days apart).

All 26 are now in `ingest_queue` as `amount_variance` with both row ids in `candidate_ids` — **queued, not deleted**. They are `funded_by='owner'` and `cash_source='personal'`, so removing them would NOT move the cash reserve; it would cut "what Come With owes Keith" (currently $25,410.76) by ~$906. The remaining 36 unreffed rows ($12,937.65) have no pair and are genuine site-only records.

Related: [[cw-cash-reserve-inflows-bug]], [[unified-finance-model]]
