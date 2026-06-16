-- =============================================================================
-- 038_guest_module.sql  —  Guest module: actor-graduation link + returning KPI
--
--  A) guests.actor_id — the guest→actor graduation link. A person can be both a
--     guest (attendee/mailing lens) and an actor (relationship/KPI lens); this ties
--     them. Nullable; set for relationship-people (donor/sponsor/vendor/dj/staff).
--  B) v_event_attendance_kpi — per event: attendees, new vs returning, repeat %.
--     "returning" = the guest attended an earlier-dated event too.
--  C) v_guest_kpis — aggregate guest KPIs (one row): total guests, subscribed,
--     paying guests, avg spend / paying guest, total guest spend, list growth basis.
--
-- ADDITIVE ONLY: 1 nullable column, 2 views. No money tables touched. security_invoker
-- views (admin-only via underlying RLS; anon-revoked).
-- =============================================================================
begin;

alter table public.guests add column if not exists actor_id uuid references public.actors(id) on delete set null;
comment on column public.guests.actor_id is
  'Graduation link: the actor this guest also is (relationship/KPI lens). Same human, two lenses.';
create index if not exists idx_guests_actor on public.guests(actor_id) where actor_id is not null;

-- B) returning-attendee KPI per event
create or replace view public.v_event_attendance_kpi with (security_invoker = true) as
with att as (
  select gea.guest_id, gea.event_id, e.event_date
  from public.guest_event_attendance gea
  join public.events e on e.id = gea.event_id and e.deleted_at is null
),
firstev as (select guest_id, min(event_date) as first_date from att group by guest_id)
select
  e.id as event_id, e.name, e.event_date,
  count(distinct a.guest_id)                                                       as attendees,
  count(distinct a.guest_id) filter (where a.event_date = f.first_date)            as new_attendees,
  count(distinct a.guest_id) filter (where a.event_date > f.first_date)            as returning_attendees,
  round(100.0 * count(distinct a.guest_id) filter (where a.event_date > f.first_date)
        / nullif(count(distinct a.guest_id), 0), 1)                               as repeat_pct
from public.events e
join att a on a.event_id = e.id
join firstev f on f.guest_id = a.guest_id
where e.deleted_at is null
group by e.id, e.name, e.event_date;
revoke all on public.v_event_attendance_kpi from anon;

-- C) aggregate guest KPIs
create or replace view public.v_guest_kpis with (security_invoker = true) as
select
  (select count(*) from public.guests where deleted_at is null)                              as total_guests,
  (select count(*) from public.subscribers where status = 'subscribed')                      as subscribed,
  (select count(distinct guest_id) from public.guest_event_attendance)                       as guests_with_attendance,
  (select count(*) from public.guests g where g.deleted_at is null
     and (select count(*) from public.guest_event_attendance a where a.guest_id=g.id) > 1)   as repeat_guests,
  (select coalesce(round(avg(total_spent),2),0) from public.v_guest_stats where total_spent > 0) as avg_spend_per_paying_guest,
  (select coalesce(sum(amount_spent),0) from public.guest_event_attendance)                  as total_guest_spend;
revoke all on public.v_guest_kpis from anon;

commit;

-- DOWN: drop view v_guest_kpis; drop view v_event_attendance_kpi; alter table guests drop column actor_id;
