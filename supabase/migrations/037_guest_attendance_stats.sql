-- =============================================================================
-- 037_guest_attendance_stats.sql  —  Attendee/mailing backfill support
--
--  A) guest_event_attendance — additive guest↔event link carrying amount_spent.
--     We do NOT write ticketing rows for historical attendees: ticketing feeds
--     v_event_summary.ticket_revenue, and DI#1's ticket money is already booked as a
--     reconciled income row — adding ticketing would double-count event revenue.
--     The guest's lifetime "total spent" lives here, on the guest side, leaving every
--     event's financials untouched.
--  B) v_guest_stats — live lifetime stats per guest: events attended (count + list),
--     total spent, first/last seen, subscribed?.
--
-- ADDITIVE ONLY: 1 table, 1 view, RLS, indexes. No DROP / destructive ALTER / data
-- deletion. security_invoker view (admin-only via underlying RLS; anon-revoked).
-- =============================================================================
begin;

create table if not exists public.guest_event_attendance (
  id           uuid primary key default gen_random_uuid(),
  guest_id     uuid not null references public.guests(id) on delete cascade,
  event_id     uuid not null references public.events(id) on delete cascade,
  amount_spent numeric(10,2) not null default 0,
  ticket_type  text,
  quantity     integer,
  source       text,
  purchased_at timestamptz,
  created_at   timestamptz not null default now()
);
create unique index if not exists idx_guest_event_attendance_unique on public.guest_event_attendance(guest_id, event_id);
create index if not exists idx_guest_event_attendance_event on public.guest_event_attendance(event_id);
alter table public.guest_event_attendance enable row level security;
create policy "Admins can manage guest attendance" on public.guest_event_attendance for all using (public.is_admin());

create or replace view public.v_guest_stats with (security_invoker = true) as
select
  g.id            as guest_id,
  g.full_name,
  g.email,
  g.opted_in_mailing,
  count(distinct gea.event_id)                                        as events_attended,
  coalesce(array_agg(distinct e.name) filter (where e.name is not null), '{}') as events,
  coalesce(sum(gea.amount_spent), 0)                                  as total_spent,
  min(gea.purchased_at)                                               as first_seen,
  max(gea.purchased_at)                                               as last_seen,
  exists (select 1 from public.subscribers s
           where lower(s.email) = lower(g.email) and s.status = 'subscribed') as subscribed
from public.guests g
left join public.guest_event_attendance gea on gea.guest_id = g.id
left join public.events e on e.id = gea.event_id and e.deleted_at is null
where g.deleted_at is null
group by g.id, g.full_name, g.email, g.opted_in_mailing;
revoke all on public.v_guest_stats from anon;

commit;

-- DOWN: drop view v_guest_stats; drop table guest_event_attendance;
