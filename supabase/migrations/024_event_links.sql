-- =============================================================================
-- 024_event_links.sql  —  Phase B: event links (additive)
-- Spec §1.2, §1.3, §1.5, §6 Phase B. NOT APPLIED — review before apply. Push held.
--
-- - event_participants: people↔events (who played/painted/crewed), role is free
--   text (endlessly extensible per spec §1.3). Backfilled from artist_bookings.
-- - events.type (canonical 4-value axis) migrated from series; is_content_event.
-- - equipment_usage: add `purpose` so the Log Event panel can record gear role
--   (UI wiring is a follow-up — schema made ready here).
-- - DORMANT external-actor RLS tier: actor-self SELECT policies + current_actor_id().
--   These grant nothing until a login is linked to a non-admin actor. NO external
--   login is provisioned (see ROADMAP blocker: financial-view lockdown gates that).
-- =============================================================================

-- Helper: the actor_id linked to the current auth user (security definer so the
-- policy can resolve it regardless of the caller's own RLS). Returns NULL for
-- admins/anon with no linked actor — so actor-self policies simply match nothing.
create or replace function public.current_actor_id() returns uuid
  language sql stable security definer set search_path = public as $$
  select id from public.actors where user_id = auth.uid() and deleted_at is null limit 1;
$$;

-- ----------------------------------------------------------------------------
-- event_participants
-- ----------------------------------------------------------------------------
create table public.event_participants (
  id            uuid primary key default gen_random_uuid(),
  event_id      uuid not null references public.events(id) on delete cascade,
  actor_id      uuid not null references public.actors(id) on delete cascade,
  role          text not null,             -- headliner|dj|opener|painter|dancer|performer|host|crew|photographer|producer... (free text)
  bill_order    integer,
  set_start     timestamptz,
  set_end       timestamptz,
  fee           numeric(10,2),             -- what they were paid; reconciled to expenses (Q3: NOT auto-created)
  is_contractor boolean not null default false,
  notes         text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);
create unique index idx_event_participants_unique on public.event_participants(event_id, actor_id, role);
create index idx_event_participants_event on public.event_participants(event_id);
create index idx_event_participants_actor on public.event_participants(actor_id);

create trigger set_updated_at before update on public.event_participants
  for each row execute function public.handle_updated_at();

alter table public.event_participants enable row level security;
create policy "Admins can manage event participants" on public.event_participants for all using (public.is_admin());
-- DORMANT tier: an actor sees ONLY their own participation rows (incl. their own fee).
create policy "Actors can read own participation" on public.event_participants
  for select using (actor_id = public.current_actor_id());

-- Backfill from the existing artist_bookings (artist×event×role×fee).
insert into public.event_participants (event_id, actor_id, role, set_start, set_end, fee, is_contractor)
select ab.event_id, sl.actor_id, coalesce(nullif(ab.role,''), 'artist'),
       ab.set_start, ab.set_end, ab.fee, false
  from public.artist_bookings ab
  join public.actor_source_links sl on sl.source_table = 'artist' and sl.source_id = ab.artist_id
on conflict (event_id, actor_id, role) do nothing;

-- ----------------------------------------------------------------------------
-- Dormant actor-self read on actors + actor_roles (the rest of the tier)
-- ----------------------------------------------------------------------------
create policy "Actors can read own actor row" on public.actors
  for select using (user_id = auth.uid());
create policy "Actors can read own roles" on public.actor_roles
  for select using (actor_id = public.current_actor_id());

-- ----------------------------------------------------------------------------
-- events.type (canonical axis) + is_content_event — migrated from series
-- ----------------------------------------------------------------------------
alter table public.events add column if not exists type text
  check (type in ('party','dance_infusion','production','showcase'));
alter table public.events add column if not exists is_content_event boolean not null default false;

update public.events set type = case
    when series = 'Come With Parties'    then 'party'
    when series = 'Dance Infusion'       then 'dance_infusion'
    when series = 'Come With Production' then 'production'
    else 'showcase'
  end
 where type is null;
update public.events set is_content_event = true where type = 'showcase' and is_content_event = false;

create index if not exists idx_events_type on public.events(type);
-- NOTE: `series` is KEPT (the KPI views match it exactly — CLAUDE.md series contract).
-- Repointing the KPI views to events.type is a later, reviewed step (not in this phase).

-- ----------------------------------------------------------------------------
-- equipment_usage: make it carry an event-gear "purpose"/role (writable from UI)
-- ----------------------------------------------------------------------------
alter table public.equipment_usage add column if not exists purpose text;
-- (RLS is_admin() already exists from 006; the Log Event panel wiring to write
--  equipment_usage is a dashboard follow-up — logged in BUILD_LOG Phase B.)

-- Grants: new table inherits 013 default privileges; only-admin via RLS + the
-- dormant actor-self SELECT. No anon grants.

-- DOWN: drop policy ...; drop table public.event_participants;
--       alter table public.events drop column type, drop column is_content_event;
--       alter table public.equipment_usage drop column purpose;
--       drop function public.current_actor_id();
