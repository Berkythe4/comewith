-- =============================================================================
-- 008_artists.sql
-- Artist directory (admin-only initially per decision #10).
-- Artists are linked to events via artist_bookings.
-- =============================================================================

create table public.artists (
  id              uuid primary key default gen_random_uuid(),
  stage_name      text not null,
  legal_name      text,
  bio             text,
  genres          text[] not null default '{}',
  signature_color text,  -- per memory: artists each have a signature color
  rate            numeric(10, 2),
  rate_unit       text,  -- 'per_set', 'per_hour', 'per_event'
  contact_email   text,
  contact_phone   text,
  social_links    jsonb not null default '{}'::jsonb,
  photo_path      text,
  logo_path       text,
  status          text not null default 'active'
                    check (status in ('active', 'inactive', 'archived')),
  reliability_score smallint,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create unique index idx_artists_stage_name on public.artists(lower(stage_name))
  where deleted_at is null;
create index idx_artists_status on public.artists(status);

create trigger set_updated_at
  before update on public.artists
  for each row execute function public.handle_updated_at();

alter table public.artists enable row level security;

-- Per decision #10: admin-only at launch. Public artist profiles are a future phase.
create policy "Admins can manage artists"
  on public.artists for all
  using (public.is_admin());

-- =============================================================================
-- Artist bookings — artist × event × role × fee
-- =============================================================================
create table public.artist_bookings (
  id              uuid primary key default gen_random_uuid(),
  artist_id       uuid not null references public.artists(id) on delete cascade,
  event_id        uuid not null references public.events(id) on delete cascade,
  role            text,
  set_start       timestamptz,
  set_end         timestamptz,
  fee             numeric(10, 2) not null default 0,
  paid            boolean not null default false,
  paid_at         timestamptz,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create unique index idx_artist_bookings_unique on public.artist_bookings(artist_id, event_id);
create index idx_artist_bookings_event_id on public.artist_bookings(event_id);
create index idx_artist_bookings_artist_id on public.artist_bookings(artist_id);

create trigger set_updated_at
  before update on public.artist_bookings
  for each row execute function public.handle_updated_at();

alter table public.artist_bookings enable row level security;

create policy "Admins can manage artist bookings"
  on public.artist_bookings for all
  using (public.is_admin());

-- =============================================================================
-- Artist notes — private, more sensitive
-- =============================================================================
create table public.artist_notes (
  id              uuid primary key default gen_random_uuid(),
  artist_id       uuid not null references public.artists(id) on delete cascade,
  note            text not null,
  author_id       uuid references public.profiles(id),
  created_at      timestamptz not null default now()
);

create index idx_artist_notes_artist_id on public.artist_notes(artist_id);

alter table public.artist_notes enable row level security;

create policy "Master admin can manage artist notes"
  on public.artist_notes for all
  using (public.is_master_admin());

create policy "Sub admins can read artist notes"
  on public.artist_notes for select
  using (public.is_admin());
