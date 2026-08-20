-- ============================================================
-- COME WITH — 182 register Data Health in the nav
--
-- The nav is data-driven from module_registry (renderNav reads it), so a screen
-- that exists in dashboard.html and not in this table is a screen nobody can
-- reach. MASTER ONLY: the findings quote amounts, payees and 1099 status, which
-- is exactly the material the 043 financial gate keeps away from sub_admins.
--
-- default_roles is deliberately EMPTY. master_only already restricts it; listing
-- roles as well would imply staff could be granted access, and the whole point of
-- the screen is that it reads across every financial view at once.
-- ============================================================
begin;

insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('data-health', '🩺 Data Health', 'Team HQ', 90, true, true, true, array[]::text[])
on conflict (key) do update
   set label = excluded.label,
       nav_group = excluded.nav_group,
       sort_order = excluded.sort_order,
       built = true,
       signed_off = true,
       master_only = true;

commit;

-- DOWN: delete from public.module_registry where key = 'data-health';
