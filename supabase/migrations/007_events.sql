-- =============================================================================
-- 007_events.sql
-- Events domain. One row per production (Dance Infusion #1, #2, #3, future CW events).
-- Linked: venues, sponsors, sponsorships, ticketing, guests, raffle_prizes.
-- =============================================================================

create table public.venues (
  id              uuid primary key default gen_random_uuid(),
  name            text not null,
  address         text,
  city            text,
  state           text,
  capacity        integer,
  contact_name    text,
  contact_email   text,
  contact_phone   text,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_venues_name on public.venues(name) where deleted_at is null;

create trigger set_updated_at
  before update on public.venues
  for each row execute function public.handle_updated_at();

alter table public.venues enable row level security;

create policy "Admins can manage venues"
  on public.venues for all
  using (public.is_admin());

-- =============================================================================
-- Events
-- =============================================================================
create table public.events (
  id              uuid primary key default gen_random_uuid(),
  slug            text not null unique,
  name            text not null,
  series          text,  -- e.g. 'Dance Infusion', 'Come With Production', etc.
  event_date      date not null,
  doors_time      time,
  end_time        time,
  venue_id        uuid references public.venues(id),
  status          text not null default 'planning'
                    check (status in ('planning', 'announced', 'on_sale', 'sold_out', 'completed', 'cancelled')),
  bar_minimum     numeric(10, 2),
  ticket_url      text,
  description     text,
  hero_image_path text,
  total_attendance integer,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_events_date on public.events(event_date desc);
create index idx_events_series on public.events(series);
create index idx_events_status on public.events(status);
create index idx_events_slug on public.events(slug);

create trigger set_updated_at
  before update on public.events
  for each row execute function public.handle_updated_at();

alter table public.events enable row level security;

create policy "Admins can manage events"
  on public.events for all
  using (public.is_admin());

create policy "Public can read non-cancelled events"
  on public.events for select
  using (status != 'cancelled' and deleted_at is null);

-- Now that events exists, add FKs that were deferred from earlier migrations.
alter table public.income
  add constraint fk_income_event foreign key (event_id) references public.events(id) on delete set null;
alter table public.expenses
  add constraint fk_expenses_event foreign key (event_id) references public.events(id) on delete set null;
alter table public.mileage
  add constraint fk_mileage_event foreign key (event_id) references public.events(id) on delete set null;
alter table public.equipment_usage
  add constraint fk_equipment_usage_event foreign key (event_id) references public.events(id) on delete set null;

-- =============================================================================
-- Sponsors + sponsorships
-- =============================================================================
create table public.sponsors (
  id              uuid primary key default gen_random_uuid(),
  name            text not null,
  contact_name    text,
  contact_email   text,
  contact_phone   text,
  website         text,
  logo_path       text,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create trigger set_updated_at
  before update on public.sponsors
  for each row execute function public.handle_updated_at();

alter table public.sponsors enable row level security;

create policy "Admins can manage sponsors"
  on public.sponsors for all
  using (public.is_admin());

create table public.sponsorships (
  id              uuid primary key default gen_random_uuid(),
  sponsor_id      uuid not null references public.sponsors(id) on delete cascade,
  event_id        uuid not null references public.events(id) on delete cascade,
  tier            text,
  cash_amount     numeric(10, 2) not null default 0,
  drink_tickets   integer not null default 0,
  entry_tickets   integer not null default 0,
  in_kind_value   numeric(10, 2) not null default 0,
  status          text not null default 'pending'
                    check (status in ('pending', 'confirmed', 'paid', 'cancelled')),
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create unique index idx_sponsorships_unique on public.sponsorships(sponsor_id, event_id);
create index idx_sponsorships_event_id on public.sponsorships(event_id);

create trigger set_updated_at
  before update on public.sponsorships
  for each row execute function public.handle_updated_at();

alter table public.sponsorships enable row level security;

create policy "Admins can manage sponsorships"
  on public.sponsorships for all
  using (public.is_admin());

-- =============================================================================
-- Guests (attendees, doubles as mailing list source)
-- =============================================================================
create table public.guests (
  id              uuid primary key default gen_random_uuid(),
  full_name       text,
  email           text,
  phone           text,
  opted_in_mailing boolean not null default false,
  source          text,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create unique index idx_guests_email_active on public.guests(lower(email))
  where deleted_at is null and email is not null;

create trigger set_updated_at
  before update on public.guests
  for each row execute function public.handle_updated_at();

alter table public.guests enable row level security;

create policy "Admins can manage guests"
  on public.guests for all
  using (public.is_admin());

-- =============================================================================
-- Ticketing — one row per ticket sold/given
-- =============================================================================
create table public.ticketing (
  id              uuid primary key default gen_random_uuid(),
  event_id        uuid not null references public.events(id) on delete cascade,
  guest_id        uuid references public.guests(id) on delete set null,
  ticket_type     text not null,
  amount_paid     numeric(10, 2) not null default 0,
  source          text,  -- 'zeffy', 'resident_advisor', 'comp', 'door'
  external_id     text,
  purchased_at    timestamptz,
  attended        boolean,
  notes           text,
  created_at      timestamptz not null default now()
);

create index idx_ticketing_event_id on public.ticketing(event_id);
create index idx_ticketing_guest_id on public.ticketing(guest_id);
create index idx_ticketing_source on public.ticketing(source);

alter table public.ticketing enable row level security;

create policy "Admins can manage ticketing"
  on public.ticketing for all
  using (public.is_admin());

-- =============================================================================
-- Raffle prizes
-- =============================================================================
create table public.raffle_prizes (
  id              uuid primary key default gen_random_uuid(),
  event_id        uuid not null references public.events(id) on delete cascade,
  prize_name      text not null,
  donor_sponsor_id uuid references public.sponsors(id) on delete set null,
  donor_name      text,
  estimated_value numeric(10, 2),
  winner_guest_id uuid references public.guests(id) on delete set null,
  winner_name     text,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create index idx_raffle_prizes_event_id on public.raffle_prizes(event_id);

create trigger set_updated_at
  before update on public.raffle_prizes
  for each row execute function public.handle_updated_at();

alter table public.raffle_prizes enable row level security;

create policy "Admins can manage raffle prizes"
  on public.raffle_prizes for all
  using (public.is_admin());

-- =============================================================================
-- Third-party donations (e.g. Crossroads — money that didn't flow through Bluevine)
-- =============================================================================
create table public.third_party_donations (
  id              uuid primary key default gen_random_uuid(),
  event_id        uuid references public.events(id) on delete set null,
  donor_name      text,
  amount          numeric(10, 2) not null,
  payment_processor text,  -- 'Crossroads', 'Venmo charity', etc.
  date            date,
  notes           text,
  created_at      timestamptz not null default now()
);

create index idx_third_party_donations_event_id on public.third_party_donations(event_id);

alter table public.third_party_donations enable row level security;

create policy "Admins can manage third-party donations"
  on public.third_party_donations for all
  using (public.is_admin());
