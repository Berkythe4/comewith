-- =============================================================================
-- 006_equipment.sql
-- Equipment inventory + per-event usage log. Powers per-use ROI analysis.
-- =============================================================================

create table public.equipment_inventory (
  id              uuid primary key default gen_random_uuid(),
  name            text not null,
  category        text,
  brand           text,
  model           text,
  serial_number   text,
  purchase_date   date,
  purchase_price  numeric(10, 2),
  current_value   numeric(10, 2),
  daily_rate      numeric(10, 2),
  status          text not null default 'available'
                    check (status in ('available', 'rented', 'maintenance', 'retired')),
  photo_path      text,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_equipment_status on public.equipment_inventory(status) where deleted_at is null;
create index idx_equipment_category on public.equipment_inventory(category) where deleted_at is null;

create trigger set_updated_at
  before update on public.equipment_inventory
  for each row execute function public.handle_updated_at();

alter table public.equipment_inventory enable row level security;

create policy "Admins can manage equipment"
  on public.equipment_inventory for all
  using (public.is_admin());

-- Public can read available equipment (for the public equipment list page).
create policy "Public can read available equipment"
  on public.equipment_inventory for select
  using (status = 'available' and deleted_at is null);

-- =============================================================================
-- Equipment usage — one row per equipment x event x date range
-- =============================================================================
create table public.equipment_usage (
  id              uuid primary key default gen_random_uuid(),
  equipment_id    uuid not null references public.equipment_inventory(id) on delete cascade,
  agreement_id    uuid references public.agreements(id) on delete set null,
  event_id        uuid,  -- FK added below after events table is created
  start_date      timestamptz not null,
  end_date        timestamptz,
  revenue_attributed numeric(10, 2),
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create index idx_equipment_usage_equipment_id on public.equipment_usage(equipment_id);
create index idx_equipment_usage_agreement_id on public.equipment_usage(agreement_id);
create index idx_equipment_usage_dates on public.equipment_usage(start_date, end_date);

create trigger set_updated_at
  before update on public.equipment_usage
  for each row execute function public.handle_updated_at();

alter table public.equipment_usage enable row level security;

create policy "Admins can manage equipment usage"
  on public.equipment_usage for all
  using (public.is_admin());
