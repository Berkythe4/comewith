-- =============================================================================
-- 104_partners_role.sql
--
-- New staff_role: 'partners' — Henry, Martin and Janelle. Partners see every
-- module EXCEPT the financial/admin ones.
--
-- "Except financials" is implemented as the boundary that already exists in this
-- schema: `module_registry.master_only`. Those four modules — income, expenses,
-- strategy (KPI money) and team (Users/permissions) — are unreachable by ANY
-- non-master account, because user_can_access_module() checks `not m.master_only`
-- before it ever looks at roles or grants. So partners are added to the default
-- roles of every non-master_only module and nothing else; the financial wall is
-- untouched and still enforced in one place.
--
-- Event-level money is separately gated by the two-flag release model (041/042/043)
-- and is NOT affected by this migration.
--
-- Also: 'marketing' loses Artist Radio (ra-market), per decision 2026-07-22.
--
-- NOT DONE HERE ON PURPOSE: ra-market.signed_off stays FALSE. Keith signs modules
-- off himself; this migration only decides WHO would see it once he does. Partners
-- are added to its default_roles so it appears for them the moment he signs off,
-- with no follow-up migration needed.
-- =============================================================================
begin;

-- (1) Allow the new value ------------------------------------------------------
alter table public.profiles drop constraint if exists profiles_staff_role_check;
alter table public.profiles add constraint profiles_staff_role_check
  check (staff_role = any (array['operations'::text, 'marketing'::text, 'full'::text, 'partners'::text]));

-- (2) Partners see every non-master_only module --------------------------------
-- Idempotent: only appends where it isn't already present.
update public.module_registry
   set default_roles = array_append(default_roles, 'partners')
 where not master_only
   and not ('partners' = any (default_roles));

-- (3) Marketing loses Artist Radio --------------------------------------------
update public.module_registry
   set default_roles = array_remove(default_roles, 'marketing')
 where key = 'ra-market';

-- (4) Assign the three partners. Liz stays 'operations' (unchanged, as asked). --
update public.profiles set staff_role = 'partners'
 where full_name in ('Henry Zaradich', 'Martin', 'Janelle Sochet')
   and role = 'sub_admin';

commit;
-- POST: staff_role accepts 'partners'; Henry/Martin/Janelle are partners; every
-- non-master_only module lists 'partners' in default_roles; ra-market no longer
-- lists 'marketing' and REMAINS signed_off = false (Keith signs it off).
-- income / expenses / strategy / team stay master_only → still master-admin only.
