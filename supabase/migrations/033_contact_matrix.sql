-- =============================================================================
-- 033_contact_matrix.sql  —  Venue/Vendor module Part 3a: the contact matrix
--
-- The missing relationship layer: who you deal with at a venue / vendor. Contacts
-- are ACTORS (reuse the actor model — no parallel people table). Two additive link
-- tables + two read views establish the matrix that accumulates as you work, so the
-- "last time you worked here" lookup can surface the right person next time.
--
--   venue_contacts  : venue_id ↔ actor_id (the contact person) + function + primary
--   vendor_contacts : vendor_actor_id (the vendor org/actor, role=vendor)
--                     ↔ contact_actor_id (a person) + function + primary
--
-- Functions are free text (suggested vocab: booking, sound, day_of, gm, security,
-- other) so the list stays editable without a migration.
--
-- Views are security_invoker=true (actor-details convention): admin-only via the
-- underlying RLS, anon-revoked. They carry a `last_event_with` recency column — the
-- SEAM for future frequency/recency ranking (v1 just orders by it; see lookup TODO).
--
-- ADDITIVE ONLY: 2 tables, 2 views, RLS policies, indexes. No DROP / destructive
-- ALTER / data deletion. Legacy venues.contact_* single-contact fields are untouched.
-- =============================================================================
begin;

-- ── venue_contacts ───────────────────────────────────────────────────────────
create table if not exists public.venue_contacts (
  id          uuid primary key default gen_random_uuid(),
  venue_id    uuid not null references public.venues(id) on delete cascade,
  actor_id    uuid not null references public.actors(id) on delete cascade,
  function    text,                       -- booking | sound | day_of | gm | security | other (free text)
  is_primary  boolean not null default false,
  notes       text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);
create unique index if not exists idx_venue_contacts_unique on public.venue_contacts(venue_id, actor_id, coalesce(function,''));
create index if not exists idx_venue_contacts_venue on public.venue_contacts(venue_id);
create trigger set_updated_at before update on public.venue_contacts
  for each row execute function public.handle_updated_at();
alter table public.venue_contacts enable row level security;
create policy "Admins can manage venue contacts" on public.venue_contacts for all using (public.is_admin());

-- ── vendor_contacts ──────────────────────────────────────────────────────────
-- A "vendor" is an actor with role=vendor (locked actor-is-vendor decision). Its
-- people-contacts are other actors. This links the vendor-actor to its people.
create table if not exists public.vendor_contacts (
  id               uuid primary key default gen_random_uuid(),
  vendor_actor_id  uuid not null references public.actors(id) on delete cascade,
  contact_actor_id uuid not null references public.actors(id) on delete cascade,
  function         text,
  is_primary       boolean not null default false,
  notes            text,
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now()
);
create unique index if not exists idx_vendor_contacts_unique on public.vendor_contacts(vendor_actor_id, contact_actor_id, coalesce(function,''));
create index if not exists idx_vendor_contacts_vendor on public.vendor_contacts(vendor_actor_id);
create trigger set_updated_at before update on public.vendor_contacts
  for each row execute function public.handle_updated_at();
alter table public.vendor_contacts enable row level security;
create policy "Admins can manage vendor contacts" on public.vendor_contacts for all using (public.is_admin());

-- ── v_venue_contacts ─────────────────────────────────────────────────────────
-- Canonical "who do we know at this venue" read. last_event_with = the most recent
-- event AT THIS VENUE where this contact actually participated — the recency signal
-- the "last time" lookup orders on, and the SEAM a future ranker would frequency-weight.
create or replace view public.v_venue_contacts with (security_invoker = true) as
select
  vc.id, vc.venue_id, vc.actor_id, vc.function, vc.is_primary, vc.notes, vc.updated_at as last_touch,
  a.display_name, a.email, a.phone, a.instagram,
  (select max(e.event_date)
     from public.events e
     join public.event_participants ep on ep.event_id = e.id
    where e.venue_id = vc.venue_id and ep.actor_id = vc.actor_id and e.deleted_at is null) as last_event_with
from public.venue_contacts vc
join public.actors a on a.id = vc.actor_id and a.deleted_at is null;
revoke all on public.v_venue_contacts from anon;

-- ── v_vendor_contacts ────────────────────────────────────────────────────────
create or replace view public.v_vendor_contacts with (security_invoker = true) as
select
  vc.id, vc.vendor_actor_id, vc.contact_actor_id, vc.function, vc.is_primary, vc.notes, vc.updated_at as last_touch,
  a.display_name, a.email, a.phone, a.instagram,
  va.display_name as vendor_name
from public.vendor_contacts vc
join public.actors a  on a.id = vc.contact_actor_id and a.deleted_at is null
join public.actors va on va.id = vc.vendor_actor_id;
revoke all on public.v_vendor_contacts from anon;

commit;

-- =============================================================================
-- DOWN (manual):
--   drop view if exists public.v_vendor_contacts;
--   drop view if exists public.v_venue_contacts;
--   drop table if exists public.vendor_contacts;
--   drop table if exists public.venue_contacts;
-- =============================================================================
