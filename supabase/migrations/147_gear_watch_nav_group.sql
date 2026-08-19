-- =============================================================================
-- 147_gear_watch_nav_group.sql
-- Fix: 146 registered the Gear Watch module under nav_group 'Operations', which
-- does not exist in this dashboard.
--
-- The sidebar is rendered by renderNav() over a FIXED list in dashboard.html:
--   NAV_GROUP_ORDER = ['Pinned','Workflow','Finance','Marketing','Venues',
--                      'Artists','Radio','Team HQ']
-- and the loop `for (const group of NAV_GROUP_ORDER)` simply never visits a
-- group that isn't in that array. A module filed under any other name is
-- registered, permitted, and completely unreachable — no button is drawn, and
-- nothing errors. The panel and its loader were live and correct the whole time.
--
-- Filed under 'Venues' because that group already holds `equipment`, and this is
-- the equipment that was stolen. sort_order 65 puts it directly after it.
--
-- A new nav group is a CODE change (that array) plus a data change. Prefer an
-- existing group unless the code lands in the same push.
-- =============================================================================
begin;

update public.module_registry
   set nav_group  = 'Venues',
       sort_order = 65
 where key = 'gearwatch';

notify pgrst, 'reload schema';
commit;

-- DOWN: update public.module_registry set nav_group='Operations' where key='gearwatch';
