-- =============================================================================
-- 003_clients_contractors.sql
-- Master client + contractor records. Clients are dedupe targets for inquiries
-- and agreements. Contractors are paid collaborators (DJs, photographers).
-- =============================================================================

create table public.clients (
  id              uuid primary key default gen_random_uuid(),
  user_id         uuid references public.profiles(id) on delete set null,
  full_name       text not null,
  email           text,
  phone           text,
  company         text,
  address         text,
  notes           text,
  source          text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create unique index idx_clients_email_active on public.clients(lower(email))
  where deleted_at is null and email is not null;
create index idx_clients_user_id on public.clients(user_id);
create index idx_clients_created_at on public.clients(created_at desc);

create trigger set_updated_at
  before update on public.clients
  for each row execute function public.handle_updated_at();

alter table public.clients enable row level security;

create policy "Admins can manage clients"
  on public.clients for all
  using (public.is_admin());

create policy "Customers can read own client record"
  on public.clients for select
  using (auth.uid() = user_id);

-- =============================================================================
-- Contractors
-- =============================================================================
create table public.contractors (
  id              uuid primary key default gen_random_uuid(),
  full_name       text not null,
  stage_name      text,
  email           text,
  phone           text,
  role            text,
  hourly_rate     numeric(10, 2),
  day_rate        numeric(10, 2),
  payment_terms   text,
  tax_form_on_file boolean not null default false,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_contractors_role on public.contractors(role) where deleted_at is null;
create index idx_contractors_email on public.contractors(lower(email)) where deleted_at is null;

create trigger set_updated_at
  before update on public.contractors
  for each row execute function public.handle_updated_at();

alter table public.contractors enable row level security;

create policy "Admins can manage contractors"
  on public.contractors for all
  using (public.is_admin());
