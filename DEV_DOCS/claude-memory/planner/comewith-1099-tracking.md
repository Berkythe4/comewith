---
name: comewith-1099-tracking
description: 1099 reportability is a stored human decision on actors.tax_1099_status, never inferred from expense category
metadata:
  type: project
---

`v_contractor_1099` (migration 166) lists Come With payees by **payee per calendar
year across all categories**, not by category. Reportability is a stored decision
in `actors.tax_1099_status` ('due' / 'exempt' / null = undecided), set by a person.

**Why:** the first version (165) grouped on `category='Contractors'` and
under-reported real money — Janelle Sochet showed $700 against $900 actually paid,
19th & 7th Productions showed $900 against $1,800. The $600 threshold applies to
total service payments to a payee in a year regardless of internal bucket. The
ledger cannot know entity type, goods-vs-services, or reimbursement status, so
inferring the answer is what produced the wrong figures.

Related root-cause fix: migration 158 had seeded **Venmo, a payment rail, as an
actor** with a matching alias rule, collapsing every Venmo payment into one payee
(168 removed it). Watch for the same pattern with PayPal/Zelle/Cash App — none
currently exist as actors.

**How to apply:** new payees over $600 surface as `needs_review`; resolve them by
setting the flag, not by recategorising the expense. See
[[comewith-ownership-and-equity]].
