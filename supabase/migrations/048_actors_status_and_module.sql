-- =============================================================================
-- 048_actors_status_and_module.sql  (additive)
--  A) actors.status — 'active' | 'on_hold' (archive = deleted_at, as elsewhere).
--     Powers the new Actors management tab's "put on hold" control.
--  B) register the 'actors' module so the data-driven nav shows the tab.
-- Admin-only via existing actors RLS; no anon grant.
-- =============================================================================
begin;

alter table public.actors
  add column if not exists status text not null default 'active'
  check (status in ('active', 'on_hold'));
create index if not exists idx_actors_status on public.actors(status) where deleted_at is null;

insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('actors', 'Actors', 'Partners', 95, true, true, false, array['operations', 'full'])
on conflict (key) do update set built = true, signed_off = true;

commit;

-- DOWN: delete from module_registry where key='actors'; alter table actors drop column status;
