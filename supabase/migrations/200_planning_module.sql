-- ============================================================
-- COME WITH — 200 register the Planning board in the module registry
--
-- Same pattern as 190 (Invoices): dashboard.html registers the tab client-side
-- so the nav does not depend on a migration having propagated, and this is the
-- other half — the registry row, so per-user module access (041/042) has
-- something to point at.
--
-- nav_group 'Finance', sort_order 16: immediately after the P&L (15). The P&L
-- is what happened; Planning is what is expected to happen. They are the same
-- statement read in two directions and belong next to each other.
-- ============================================================
begin;

insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only)
values ('planning', '📈 Planning', 'Finance', 16, true, false, true)
on conflict (key) do update
  set label       = excluded.label,
      nav_group   = excluded.nav_group,
      sort_order  = excluded.sort_order,
      built       = excluded.built,
      master_only = excluded.master_only;

commit;

-- DOWN: delete from public.module_registry where key = 'planning';
