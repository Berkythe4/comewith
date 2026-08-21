-- ============================================================
-- COME WITH — 190 register Invoices in the module registry
--
-- 188 built the module and dashboard.html registers it client-side, the same
-- resilience trick data-health and scene-map use so the nav does not depend on
-- a migration having propagated. This is the other half: the registry row, so
-- per-user module access (041/042) can actually be granted or withheld for it
-- like any other module. Without a row there is nothing for an override to
-- point at.
--
-- nav_group is 'Finance', NOT a new 'Money' group. Invoices belongs beside
-- Income, the P&L and Expenses — it bills the rows Income holds. sort_order 12
-- puts it between Income (10) and the P&L (15), which is the order the work
-- actually happens in: record the revenue, bill it, report it.
-- ============================================================
begin;

insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only)
values ('invoices', '🧾 Invoices', 'Finance', 12, true, true, false)
on conflict (key) do update
  set label = excluded.label,
      nav_group = excluded.nav_group,
      sort_order = excluded.sort_order,
      built = excluded.built,
      signed_off = excluded.signed_off;

commit;

-- DOWN: delete from public.module_registry where key = 'invoices';
