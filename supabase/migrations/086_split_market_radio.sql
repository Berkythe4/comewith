-- =============================================================================
-- 086_split_market_radio.sql
-- Split the RA Market screen into two modules: "Best Nights" (scheduling intel —
-- best nights, calendar, venues, genres) and "Artist Radio" (the existing
-- ra-market panel, retitled). Both Insights, master-only until signed off.
-- =============================================================================
begin;
update public.module_registry set label = 'Artist Radio', sort_order = 197 where key = 'ra-market';
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('market', 'Best Nights', 'Insights', 196, true, false, false, '{marketing,full}')
on conflict (key) do update set label = excluded.label, sort_order = excluded.sort_order;
commit;
-- POST: nav shows "Best Nights" (market) + "Artist Radio" (ra-market).
