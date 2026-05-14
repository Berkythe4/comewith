-- =============================================================================
-- 005_financials.sql
-- Income ledger, expense ledger, and mileage log. All admin-only.
-- =============================================================================

create table public.income (
  id              uuid primary key default gen_random_uuid(),
  date            date not null,
  client_id       uuid references public.clients(id) on delete set null,
  agreement_id    uuid references public.agreements(id) on delete set null,
  event_id        uuid,  -- FK added in 006 after events table exists
  amount          numeric(10, 2) not null,
  category        text,
  payment_method  text,
  description     text,
  receipt_path    text,
  created_by      uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_income_date on public.income(date desc);
create index idx_income_client_id on public.income(client_id);
create index idx_income_agreement_id on public.income(agreement_id);

create trigger set_updated_at
  before update on public.income
  for each row execute function public.handle_updated_at();

alter table public.income enable row level security;

create policy "Admins can manage income"
  on public.income for all
  using (public.is_admin());

-- =============================================================================
-- Expenses
-- =============================================================================
create table public.expenses (
  id              uuid primary key default gen_random_uuid(),
  date            date not null,
  event_id        uuid,  -- FK added in 006
  amount          numeric(10, 2) not null,
  category        text,
  vendor          text,
  payment_method  text,
  description     text,
  receipt_path    text,
  deductible      boolean not null default true,
  created_by      uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_expenses_date on public.expenses(date desc);
create index idx_expenses_category on public.expenses(category);
create index idx_expenses_event_id on public.expenses(event_id);

create trigger set_updated_at
  before update on public.expenses
  for each row execute function public.handle_updated_at();

alter table public.expenses enable row level security;

create policy "Admins can manage expenses"
  on public.expenses for all
  using (public.is_admin());

-- =============================================================================
-- Mileage log
-- =============================================================================
create table public.mileage (
  id              uuid primary key default gen_random_uuid(),
  date            date not null,
  client_id       uuid references public.clients(id) on delete set null,
  event_id        uuid,  -- FK added in 006
  origin          text,
  destination     text,
  miles           numeric(8, 2) not null,
  purpose         text,
  notes           text,
  created_by      uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_mileage_date on public.mileage(date desc);

create trigger set_updated_at
  before update on public.mileage
  for each row execute function public.handle_updated_at();

alter table public.mileage enable row level security;

create policy "Admins can manage mileage"
  on public.mileage for all
  using (public.is_admin());
