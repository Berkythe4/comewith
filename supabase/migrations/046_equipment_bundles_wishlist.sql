-- =============================================================================
-- 046_equipment_bundles_wishlist.sql  (additive, reversible)
--
--  A) equipment_inventory.status gains 'wishlist' (desired-but-not-owned gear, so
--     anyone with screen access can log wants). Becomes a filter on the Equipment tab.
--  B) equipment_components — internal-ops gear bundling: a child accessory/wire/misc
--     "travels with" a parent DJ/Sound item. When the parent is assigned to an event
--     (equipment_usage), the dashboard auto-adds its children.
--
-- ADDITIVE: 1 constraint widen + 1 table. Admin-only RLS; no anon grant (013 default
-- privileges cover the new table). Not financial.
-- =============================================================================
begin;

-- A) widen status check to include 'wishlist'
alter table public.equipment_inventory drop constraint if exists equipment_inventory_status_check;
alter table public.equipment_inventory add constraint equipment_inventory_status_check
  check (status in ('available', 'rented', 'maintenance', 'retired', 'wishlist'));

-- B) gear bundling (parent gear -> child accessory/wire that travels with it)
create table public.equipment_components (
  id         uuid primary key default gen_random_uuid(),
  parent_id  uuid not null references public.equipment_inventory(id) on delete cascade,
  child_id   uuid not null references public.equipment_inventory(id) on delete cascade,
  created_at timestamptz not null default now(),
  constraint equipment_components_no_self check (parent_id <> child_id)
);
create unique index idx_equipment_components_uniq on public.equipment_components(parent_id, child_id);
create index idx_equipment_components_parent on public.equipment_components(parent_id);
create index idx_equipment_components_child on public.equipment_components(child_id);
alter table public.equipment_components enable row level security;
create policy "Admins manage equipment components"
  on public.equipment_components for all using (public.is_admin());

commit;

-- DOWN:
--   drop table public.equipment_components;
--   alter table public.equipment_inventory drop constraint equipment_inventory_status_check;
--   alter table public.equipment_inventory add constraint equipment_inventory_status_check
--     check (status in ('available','rented','maintenance','retired'));
