---
name: project-site-owner-and-admin-access
description: "Martin + Henry promoted to master_admin 2026-08-15; profiles.is_owner (Keith, one row) + protect_site_owner() trigger in migration 140 stops any other admin demoting/deactivating/deleting the owner"
metadata: 
  node_type: memory
  type: project
  originSessionId: 6264aafb-c0a6-4e5c-bdca-a6f2c2a39988
  modified: 2026-08-15T22:04:46.407Z
---

Applied to prod 2026-08-15 (migration **140_site_owner.sql**, `invite-user` v9).

**Who's what now:** `master_admin` = Keith, **Martin, Henry**. `sub_admin` = Janelle,
Liz. Still no `admin` role.

**Why a trigger and not a policy:** `master_admin` was the top of the role system, so
"full admin minus one thing" can't be a role. `"Master admin can manage all profiles"`
is `for all using (is_master_admin())` with **no WITH CHECK**, so any master_admin
could `PATCH /profiles?id=eq.<keith>` through PostgREST — and the dashboard has no
role-change control at all, so the UI was never the boundary.

`profiles.is_owner` (Keith, exactly one row) + `protect_site_owner()`, a
`before insert or update or delete` trigger. It refuses, from anyone but the owner:
changes to the owner's `role`, `deleted_at` or `is_owner`, and `DELETE` of that row.
Name/phone/staff_role stay editable. Ownership is **given** by the owner, never taken.

**`deleted_at` is the vector people miss** — under the 098 deactivation contract a
deactivated profile reads as no-role to `is_admin()` / `is_master_admin()` /
`user_can_access_module()`, so deactivating the owner locks them out while `role`
still says `master_admin`. Guard both or it isn't a guard.

**Deliberate hole:** the trigger exempts `auth.uid() is null` (service role, Edge
Functions, Management API) so break-glass repair works. It therefore protects the
**app**, not the **project** — service-role key, `SBP_PAT`, Supabase dashboard, GitHub
and Netlify all sit above it and stay Keith-only. Any service-role edge function that
writes `profiles` must re-enforce the rule; `invite-user` refuses the owner's email.

**Verified on prod** with 10 impersonation checks as Martin inside BEGIN..ROLLBACK:
5 blocked, 5 correctly allowed. Users tab shows a 👑 owner chip (`get_team_members()`
was dropped/recreated to return `is_owner` — adding a column to a RETURNS TABLE needs
a drop, not a replace).

**Accepted trade, stated to Keith before applying:** the two new masters can remove
each other, invite more masters, read/write all company money, and flip
`financials_released` — so the 041–043 financial gate now covers only Janelle and Liz.
Peer-locking the masters is NOT built. Full reasoning: LEARNINGS §19.
Related: [[project-staff-access-model]], [[project-user-management]], [[project-two-machine-handoff]].
