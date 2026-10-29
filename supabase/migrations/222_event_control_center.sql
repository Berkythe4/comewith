-- =============================================================================
-- 222_event_control_center.sql — the Event Control Center (control.html)
-- =============================================================================
-- One admin tool for every Come With event, with a fundraising layer that
-- switches on per event (Dance Infusion first: DI #3, 2026-11-14 at Signal).
--
-- Most of what the Control Center shows ALREADY HAS A HOME, and it stays there:
--
--   ticket sales          ticketing                 (+ tier_id, below)
--   the door list         guest_event_attendance    (one row per person per
--                         event: idx_guest_event_attendance_unique;
--                         + entry_type / sponsor link / drink tickets /
--                         check-in stamp, below)
--   people                guests                    (unique lower(email))
--   sponsors              sponsorships              (entry_tickets = comps,
--                         drink_tickets = the package, status pending/paid)
--   donations             third_party_donations     (Zeffy, MS Society DIY)
--   actual money          income / expenses         (the P&L; untouched)
--   planned money         budget_lines scope='event' (the hub's FORECAST
--                         column - never in the P&L)
--   blockers              tasks.status = 'blocked'  (they already exist as a
--                         status, so they show on Calendar & Tasks too)
--
-- New here, because nothing held it:
--
--   event_control            per-event switches + venue terms: is fundraising
--                            on, beneficiary, % to mission target, the drink
--                            ticket count locked with the venue, and what one
--                            drink ticket is worth toward the bar minimum
--   event_ticket_tiers       price and its split (to mission / production);
--                            the production share is DERIVED (price - to_mission),
--                            never stored, so the two cannot disagree
--   event_drink_allocations  block allocations nobody is named on yet
--                            ("DJ pool 10", "staff 20")
--   event_exec_updates       saved snapshots of the exec update, so the next one
--                            can say what closed since
--
-- Drink tickets are counted ONCE (DI #2 rule). The ledger in
-- assets/control-center.js reads:
--   sponsor packages   sponsorships.drink_tickets (handed-out tickets on a
--                      sponsor's comp rows are PART of that package, never added)
--   ticket tiers       tier.drink_tickets x tickets sold
--   people             guest_event_attendance.drink_tickets on rows with no
--                      sponsorship_id (DJs, staff, volunteers, comps)
--   blocks             event_drink_allocations
--
-- "TBD" is not zero. Two existing NOT NULL DEFAULT 0 columns gain a boolean
-- beside them rather than losing NOT NULL, because v_event_summary,
-- v_event_forecast, v_budget_variance and v_pl_monthly read those columns and
-- were written for them never being null:
--   sponsorships.drink_tickets_tbd   the package is not agreed yet
--   budget_lines.amount_tbd          a planned cost with no number yet
--
-- ADDITIVE. No column dropped, no constraint tightened on a column that holds
-- data, nothing rewritten. The deployed dashboard ignores every new column.
--
-- RLS: every new table is admin-only (is_admin(): master_admin + sub_admin,
-- deactivation-aware). Existing tables keep their policies, so money writes
-- (ticketing, sponsorships, donations, income, expenses) stay master-only and
-- money reads stay behind can_see_event_financials() - the Control Center does
-- not widen the financial gate.
-- =============================================================================

begin;

-- ---------------------------------------------------------------------------
-- event_control
-- ---------------------------------------------------------------------------
create table if not exists public.event_control (
  event_id             uuid primary key references public.events(id) on delete cascade,
  fundraising          boolean not null default false,
  beneficiary          text,
  mission_target_pct   numeric(5,2) check (mission_target_pct is null or (mission_target_pct >= 0 and mission_target_pct <= 100)),
  drink_tickets_locked integer check (drink_tickets_locked is null or drink_tickets_locked >= 0),
  drink_ticket_value   numeric(10,2) check (drink_ticket_value is null or drink_ticket_value >= 0),
  comp_mission_value   numeric(10,2) check (comp_mission_value is null or comp_mission_value >= 0),
  notes                text,
  created_at           timestamptz not null default now(),
  updated_at           timestamptz not null default now(),
  updated_by           uuid default auth.uid() references public.profiles(id) on delete set null
);
comment on table  public.event_control is 'Event Control Center: per-event switches and venue terms. No row = defaults (fundraising on for series Dance Infusion).';
comment on column public.event_control.drink_tickets_locked is 'Drink tickets locked with the venue. NULL = not locked yet (renders TBD, never 0).';
comment on column public.event_control.drink_ticket_value is '$ one drink ticket counts toward events.bar_minimum. NULL = unknown; bar-minimum progress is then not computed.';
comment on column public.event_control.comp_mission_value is '$ to the beneficiary credited per sponsor comp (DI #2 counted $25). NULL = comps credit nothing.';

-- ---------------------------------------------------------------------------
-- event_ticket_tiers
-- ---------------------------------------------------------------------------
create table if not exists public.event_ticket_tiers (
  id               uuid primary key default gen_random_uuid(),
  event_id         uuid not null references public.events(id) on delete cascade,
  label            text not null,
  price            numeric(10,2) not null check (price >= 0),
  to_mission       numeric(10,2) not null default 0 check (to_mission >= 0),
  production_label text,
  drink_tickets    integer check (drink_tickets is null or drink_tickets >= 0),
  aliases          text[] not null default '{}',
  sort_order       integer not null default 0,
  active           boolean not null default true,
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now(),
  constraint event_ticket_tiers_split check (to_mission <= price)
);
create index if not exists event_ticket_tiers_event on public.event_ticket_tiers (event_id, sort_order);
comment on column public.event_ticket_tiers.drink_tickets is 'Drink tickets included per ticket. NULL = TBD.';
comment on column public.event_ticket_tiers.aliases is 'Ticket-type strings from Zeffy/RA/DIY exports that map to this tier (lowercased).';
comment on column public.event_ticket_tiers.production_label is 'Name of the (price - to_mission) portion, e.g. "Production support". Editable.';

-- ---------------------------------------------------------------------------
-- event_drink_allocations
-- ---------------------------------------------------------------------------
create table if not exists public.event_drink_allocations (
  id         uuid primary key default gen_random_uuid(),
  event_id   uuid not null references public.events(id) on delete cascade,
  category   text not null check (category in ('dj', 'staff', 'volunteer', 'comp', 'purchased', 'other')),
  label      text,
  tickets    integer not null check (tickets >= 0),
  notes      text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists event_drink_allocations_event on public.event_drink_allocations (event_id);

-- ---------------------------------------------------------------------------
-- event_exec_updates
-- ---------------------------------------------------------------------------
create table if not exists public.event_exec_updates (
  id         uuid primary key default gen_random_uuid(),
  event_id   uuid not null references public.events(id) on delete cascade,
  snapshot   jsonb not null,
  note       text,
  created_at timestamptz not null default now(),
  created_by uuid default auth.uid() references public.profiles(id) on delete set null
);
create index if not exists event_exec_updates_event on public.event_exec_updates (event_id, created_at desc);

-- ---------------------------------------------------------------------------
-- Columns on existing tables (all nullable or defaulted)
-- ---------------------------------------------------------------------------
alter table public.ticketing
  add column if not exists tier_id uuid references public.event_ticket_tiers(id) on delete set null;
-- Imports dedupe on (source, external_id) in code (import-engine.js and
-- control-center.js both); this makes a double import impossible rather than
-- merely unlikely. Prod held 0 duplicate keys on 2026-10-29; hand-entered and
-- recon rows carry no external_id and are unaffected.
create unique index if not exists ticketing_event_source_external
  on public.ticketing (event_id, source, external_id)
  where external_id is not null;

alter table public.guest_event_attendance
  add column if not exists entry_type     text,
  add column if not exists sponsorship_id uuid references public.sponsorships(id) on delete set null,
  add column if not exists tier_id        uuid references public.event_ticket_tiers(id) on delete set null,
  add column if not exists drink_tickets  integer not null default 0,
  add column if not exists checked_in_at  timestamptz,
  add column if not exists checked_in_by  uuid references public.profiles(id) on delete set null,
  add column if not exists notes          text;

do $$ begin
  alter table public.guest_event_attendance
    add constraint gea_entry_type_check check (entry_type is null or entry_type in
      ('ticket', 'comp', 'sponsor_comp', 'guest_list', 'dj', 'artist', 'staff', 'volunteer', 'vendor', 'walk_in'));
exception when duplicate_object then null; end $$;
do $$ begin
  alter table public.guest_event_attendance
    add constraint gea_drink_tickets_check check (drink_tickets >= 0);
exception when duplicate_object then null; end $$;
create index if not exists gea_sponsorship on public.guest_event_attendance (sponsorship_id) where sponsorship_id is not null;

alter table public.sponsorships
  add column if not exists drink_tickets_tbd boolean not null default false;

alter table public.budget_lines
  add column if not exists amount_tbd boolean not null default false;

-- Donations had no stable key, so a re-import could double them. The importer
-- writes one; hand-entered rows leave it null and are unaffected.
alter table public.third_party_donations
  add column if not exists external_ref text;
create unique index if not exists third_party_donations_external_ref
  on public.third_party_donations (event_id, payment_processor, external_ref)
  where external_ref is not null;

-- ---------------------------------------------------------------------------
-- updated_at stamps
-- ---------------------------------------------------------------------------
create or replace function public.event_control_touch()
returns trigger language plpgsql set search_path = public as $$
begin
  new.updated_at := now();
  return new;
end $$;
revoke all on function public.event_control_touch() from public, anon;

drop trigger if exists event_control_touch on public.event_control;
create trigger event_control_touch before update on public.event_control
  for each row execute function public.event_control_touch();
drop trigger if exists event_ticket_tiers_touch on public.event_ticket_tiers;
create trigger event_ticket_tiers_touch before update on public.event_ticket_tiers
  for each row execute function public.event_control_touch();
drop trigger if exists event_drink_allocations_touch on public.event_drink_allocations;
create trigger event_drink_allocations_touch before update on public.event_drink_allocations
  for each row execute function public.event_control_touch();

-- ---------------------------------------------------------------------------
-- RLS: admin-only, a real policy on every table
-- ---------------------------------------------------------------------------
alter table public.event_control           enable row level security;
alter table public.event_ticket_tiers      enable row level security;
alter table public.event_drink_allocations enable row level security;
alter table public.event_exec_updates      enable row level security;

drop policy if exists "Admins manage event control" on public.event_control;
create policy "Admins manage event control" on public.event_control
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage ticket tiers" on public.event_ticket_tiers;
create policy "Admins manage ticket tiers" on public.event_ticket_tiers
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage drink allocations" on public.event_drink_allocations;
create policy "Admins manage drink allocations" on public.event_drink_allocations
  for all using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage exec updates" on public.event_exec_updates;
create policy "Admins manage exec updates" on public.event_exec_updates
  for all using (public.is_admin()) with check (public.is_admin());

-- 013's default privileges grant anon on new tables; RLS already returns
-- nothing, but these hold guest/money config, so take the grant away too.
revoke all on public.event_control, public.event_ticket_tiers,
              public.event_drink_allocations, public.event_exec_updates from anon;

notify pgrst, 'reload schema';

commit;

-- DOWN:
--   drop table public.event_exec_updates, public.event_drink_allocations;
--   alter table public.ticketing drop column tier_id;
--   alter table public.guest_event_attendance drop column entry_type, drop column sponsorship_id,
--     drop column tier_id, drop column drink_tickets, drop column checked_in_at,
--     drop column checked_in_by, drop column notes;
--   drop table public.event_ticket_tiers, public.event_control;
--   alter table public.sponsorships drop column drink_tickets_tbd;
--   alter table public.budget_lines drop column amount_tbd;
--   drop index third_party_donations_external_ref; alter table public.third_party_donations drop column external_ref;
--   drop function public.event_control_touch();
