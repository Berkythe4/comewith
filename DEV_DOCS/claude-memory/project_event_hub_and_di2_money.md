---
name: project_event_hub_and_di2_money
description: "Event-hub Files/Customers redesign (migration 045) + how DI#2 financials are modeled (reconciled, don't re-itemize)"
metadata: 
  node_type: memory
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

Event-hub polish shipped 2026-06-24 (commit 19882ca, live on Netlify): Overview "Generate task checklist" moved to bottom; **Contracts tab removed**, replaced by a unified **Files** tab grouped into doc-type "buckets" backed by the new `document_types` registry (admin-extensible via "+ Add document type"); Files now also surfaces files attached to the event's contracts (the old tab only queried `subject_type='event'`, which is why an uploaded contract appeared "lost"). New **Customers** tab = deduped union of event_participants + sponsorships + third_party_donations + guest_event_attendance.

**Migration 045 applied to prod 2026-06-24** (via Management API, [[feedback_prod_migration_apply]]): `files.vendor_actor_id` (+FK `files_vendor_actor_id_fkey`, needed to disambiguate the PostgREST embed from `files_uploaded_by_fkey`) and `document_types` table (admin-only RLS, no anon grant).

**DI#2 (event `ff2b1917-…`) money model — NOW FULLY ITEMIZED (2026-06-24, commit 5224850), ties line-by-line to the audit:** **total_raised $9,557.33, expenses $6,557.33, net to MS Society $3,000** (unchanged headline). Components in prod: ticket_revenue **$2,313** (ticketing: 5 aggregate rows — Zeffy Guest List/Drink Pkg, RA, DIY tickets, + a "Comp & sponsor entries" $0 balancing row; **tickets_sold = 117** = 81 paid qty + 36 comp/sponsor to match attendance 117; capacity 300 → 39% sell-through), donations **$1,017.44** (third_party_donations: Keith $162.44, Crossroads $130, 12 named Zeffy donors $455, "MS Society DIY page donations" $270 aggregate), sponsor_cash **$6,225** (12 sponsorships via `actor_id` not legacy `sponsor_id` — that join was why Money-tab names were blank), other_income **$1.89** bank interest. The old **$3,039.89 "consolidated" income plug was soft-deleted** (it existed only to make the total hit $9,557.33). Audit anchor: `reports/AUDIT_TABLES/revenue_summary.csv` TOTAL $9,264.89 (Bluevine/RA/DIY/interest) + $292.44 founder/Crossroads third-party = $9,557.33. Per-person ticket counts also written to `guest_event_attendance.quantity`; Customers tab has a Reconciliation block. The working ledger `events/dance-infusion/di-02-2026-05/data/dance_infusion_ledger.csv` records GROSS (sponsorships appear twice = $9,950; Jennifer Taveras $60 zeffe_pkg is a DIY donation not a ticket) — reconcile to the AUDIT tables, not raw ledger sums.
