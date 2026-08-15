---
name: project_staff_access_model
description: Staff access model — 041+042+043 ALL applied to prod; module RLS + two-flag financial gate live; martin/henry operations logins created (2026-06-25)
metadata: 
  node_type: memory
  type: project
  originSessionId: 8c7b42a9-8c6e-4ec8-8ee5-50c34f2040a4
---

Staff access / module-gate system for dashboard.html. Built 2026-06-23 to prep
logins for Martin (operations), Henry (operations), Janelle (marketing) — accounts
NOT yet created (scaffold only, per Keith).

**Prod state:** 1 master_admin (berky@comewith.org), 1 sub_admin (liz@comewith.org).

**Migration 041 (APPLIED to prod 2026-06-23, additive/zero-risk):**
- `profiles.staff_role` ∈ (operations|marketing|full); existing sub_admins backfilled
  to 'full' (liz) so the nav gate never empties them. master_admin stays NULL.
- `module_registry` (19 rows): key=dashboard data-tab, nav_group, sort_order, built,
  signed_off, master_only, default_roles[]. **Only Events + Team are signed_off.**
  Finance/Strategy/Team are master_only. Social Calendar is the only `built=false`.
- `user_module_access` (per-person grant/revoke overrides).
- `user_can_access_module(key)` SQL fn. RLS: admins read registry, master writes;
  users read own overrides, master writes.

**Gate logic (mirrored in JS `canAccessModule` + SQL):** non-master sees a module iff
`built && signed_off && !master_only && (staff_role ∈ default_roles OR per-user grant)
&& not per-user revoke`. signed_off is ABSOLUTE — a grant can't reveal an unsigned
module. master_admin sees everything. staff_role (scope) and signed_off (release) are
orthogonal axes.

**042 + 043 APPLIED to prod 2026-06-25 (both REWRITTEN first; commit a668661):**
- **042** = hard per-module RLS. Rewrote it for the post-047 actor model — the original
  draft targeted the dropped clients/sponsors/artists/artist_bookings (would have errored).
  Now gates `actors`/`actor_roles`/`event_participants` (via new `can_see_people()` =
  actors|clients|sponsors|artists|events) + contracts/files/document_types; [VERIFY] names
  resolved against live pg_policies. Events-hub carve `can_use_events_module()` intact. Master
  full access (user_can_access_module returns true for master). 20 module policies; no leftover
  "Admins can manage" on gated tables.
- **043** = **two-flag** financial gate (Keith's final model): `events.audited` (master-only,
  informational, drives warning severity) + `events.financials_released` (master-only) — staff
  see an event's money ONLY when released; audit is NOT a hard prerequisite, just a louder
  warning. Base-table RLS on income/expenses/mileage/ticketing/sponsorships/third_party_donations
  + money cols CASE-gated in v_event_summary + `security_invoker` on all 6 money views via
  **ALTER VIEW** (so 051's v_kpi_computed/v_kpi_dashboard are NOT dropped — avoids cascade) +
  anon-revoke re-asserted. `guard_event_finance_flags` trigger blocks non-master from flipping
  either flag (INSERT+UPDATE). `can_see_event_financials(event_id)` = master OR released; null-
  event_id (company) rows stay master-only permanently. income/expenses WRITES master-only (D1);
  ticketing/sponsorships/donations writes via events module (D2). See [[feedback_no_broad_anon_grants]].
- **Dashboard UI:** master-only "Financial visibility" block in the event-hub overview
  (`hubFinVisHTML`/`hubSetAudited`/`hubToggleRelease`/`hubDoRelease`, gated by `ME.isMaster`);
  release pops a confirm EVERY time — loud red modal when not audited, gentle when audited.
- **Logins created:** invite-user deployed; **martin@comewith.org + henry@comewith.org** created
  via admin API (sub_admin / operations, email_confirm). Passwords were handed to Keith in chat
  (not stored). Janelle (marketing) not yet created. liz still sub_admin/full.
- **Gate verified end-to-end 17/17** (seeded demo event, since cleared): anon→401 ×6 views;
  master sees money pre-release; staff blocked pre-release (0 income/expense rows, NULL view
  money, 0 company rows); staff release PATCH→guard 400; master release→staff sees that event's
  money but never company finance. Test harness: martin REST JWT vs anon vs master-context SQL
  (`set_config('request.jwt.claims', sub=keith)` makes is_master_admin() true for view reads).
- **Residual** (ROADMAP): v_budget_variance / v_data_points / mv_event_data_points still need
  locking (MV can't use RLS). → CLOSED by **055** (revoked from anon+authenticated).
- **055** (2026-06-25): locked the 3 analytics views/MV from anon+authenticated.
- **057 — LEAK FIX** (2026-06-25): 043's ticketing/sponsorships/third_party_donations WRITE policy
  was `for all using can_use_events_module()`; `for all` covers SELECT, so events-module staff could
  read those revenue tables via REST. Repointed all three writes to `is_master_admin()` (matching
  income/expenses). Re-verified martin (operations) now gets **0 rows on all SIX financial tables**
  (income/expenses/mileage/ticketing/sponsorships/donations) + NULL money in views. **Lesson:** a
  `for all` write policy also grants SELECT — never pair a permissive `for all` write with a
  separately-gated read on the same table; use master-only `for all`, or split write into ins/upd/del.
- **MONEY-ADJACENT still visible to operations staff** (have Events+Equipment access, not yet decided):
  equipment_inventory.purchase_price/current_value/daily_rate; contracts.fee; event_participants.fee;
  agreements.total_amount (Agreements tab not signed off but table readable via events carve). These
  are NOT in the gated P&L ledger — open decision whether to gate them too.
- **Modules operations staff currently SEE** (signed_off + operations in default_roles): Events,
  Venues, Equipment, Actors, Sponsors, Sponsorships, Artists, Notes, Conversations. (Sponsors/
  Sponsorships tabs render names but $0 — sponsorships gated. Guests/Subscribers/Campaigns/Social
  Calendar signed off but NOT in operations role. Inquiries/Agreements/Clients/Templates not signed off.)

**Deployed:** merged to master + live on comewith.org via Netlify 2026-06-23
(merge commit 8fa2a62). Keith chose to deploy knowingly while liz@comewith.org
(sub_admin, login-capable, never signed in) still exists — since 042/043 are
unapplied, the nav gate is cosmetic and liz would retain full data access via API
if she logged in.

**dashboard.html changes (data-driven nav):** nav is now data-driven via
`renderNav()` from module_registry — do NOT re-add static nav buttons. Groups:
Sales/Operations/Finance/Partners/Audience/Insights. Master-only signed-off badges.
New panels: `#panel-social-calendar` (placeholder) + `#panel-team` (master-only Team
mgmt: sign off modules, set staff_role, cycle per-user override chips). The client-side
gate is the ONLY live enforcement until 042/043 are applied.
