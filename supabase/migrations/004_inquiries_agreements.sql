-- =============================================================================
-- 004_inquiries_agreements.sql
-- Inquiries (public form submissions) → Agreements (signed contracts).
-- agreement_links table generates prefilled URLs for clients to sign.
-- =============================================================================

create table public.inquiries (
  id              uuid primary key default gen_random_uuid(),
  client_id       uuid references public.clients(id) on delete set null,
  full_name       text not null,
  email           text not null,
  phone           text,
  event_type      text,
  event_date      date,
  venue           text,
  services_selected jsonb not null default '[]'::jsonb,
  message         text,
  source          text default 'website',
  status          text not null default 'new'
                    check (status in ('new', 'contacted', 'quoted', 'converted', 'lost', 'archived')),
  assigned_to     uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_inquiries_status on public.inquiries(status) where deleted_at is null;
create index idx_inquiries_event_date on public.inquiries(event_date);
create index idx_inquiries_created_at on public.inquiries(created_at desc);
create index idx_inquiries_email on public.inquiries(lower(email));

create trigger set_updated_at
  before update on public.inquiries
  for each row execute function public.handle_updated_at();

alter table public.inquiries enable row level security;

-- Public can INSERT (this is how the website form submits).
create policy "Anyone can submit an inquiry"
  on public.inquiries for insert
  with check (true);

create policy "Admins can read all inquiries"
  on public.inquiries for select
  using (public.is_admin());

create policy "Admins can update inquiries"
  on public.inquiries for update
  using (public.is_admin());

create policy "Master admin can delete inquiries"
  on public.inquiries for delete
  using (public.is_master_admin());

-- =============================================================================
-- Agreements (events services + equipment rental, unified)
-- =============================================================================
create table public.agreements (
  id              uuid primary key default gen_random_uuid(),
  inquiry_id      uuid references public.inquiries(id) on delete set null,
  client_id       uuid references public.clients(id) on delete set null,
  agreement_type  text not null check (agreement_type in ('events', 'rental')),
  status          text not null default 'draft'
                    check (status in ('draft', 'sent', 'signed', 'cancelled', 'completed')),

  -- Event details
  event_date      date,
  event_start_time time,
  event_end_time  time,
  venue_name      text,
  venue_address   text,

  -- Service / equipment specifics (flexible JSONB)
  services        jsonb not null default '[]'::jsonb,
  equipment       jsonb not null default '[]'::jsonb,

  -- Financials
  subtotal        numeric(10, 2) not null default 0,
  deposit_amount  numeric(10, 2) not null default 0,
  total_amount    numeric(10, 2) not null default 0,
  payment_method  text,
  payment_notes   text,
  recording_rights text,
  promo_rights    text,

  -- Rental specific
  rental_start    timestamptz,
  rental_return   timestamptz,

  -- PDF + signatures
  signed_pdf_path text,
  client_signature_url text,
  client_signed_at timestamptz,
  admin_signature_url text,
  admin_signed_at timestamptz,
  notes           text,

  created_by      uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_agreements_status on public.agreements(status) where deleted_at is null;
create index idx_agreements_type on public.agreements(agreement_type) where deleted_at is null;
create index idx_agreements_event_date on public.agreements(event_date);
create index idx_agreements_inquiry_id on public.agreements(inquiry_id);
create index idx_agreements_client_id on public.agreements(client_id);

create trigger set_updated_at
  before update on public.agreements
  for each row execute function public.handle_updated_at();

alter table public.agreements enable row level security;

create policy "Admins can manage all agreements"
  on public.agreements for all
  using (public.is_admin());

create policy "Customers can read own agreements"
  on public.agreements for select
  using (
    client_id in (select id from public.clients where user_id = auth.uid())
  );

-- =============================================================================
-- Agreement links: prefilled URLs sent to clients for review/signature.
-- Token-based, so they don't require login.
-- =============================================================================
create table public.agreement_links (
  id              uuid primary key default gen_random_uuid(),
  agreement_id    uuid not null references public.agreements(id) on delete cascade,
  token           text not null unique default encode(gen_random_bytes(24), 'hex'),
  expires_at      timestamptz not null default (now() + interval '30 days'),
  used_at         timestamptz,
  created_at      timestamptz not null default now()
);

create index idx_agreement_links_token on public.agreement_links(token);
create index idx_agreement_links_agreement_id on public.agreement_links(agreement_id);

alter table public.agreement_links enable row level security;

create policy "Admins can manage agreement links"
  on public.agreement_links for all
  using (public.is_admin());

-- Public read by token (the link itself proves authorization).
-- The serving Edge Function validates expires_at + used_at.
create policy "Anyone can read agreement link by token"
  on public.agreement_links for select
  using (true);
