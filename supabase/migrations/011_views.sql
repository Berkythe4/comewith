-- =============================================================================
-- 011_views.sql
-- Computed views for dashboards + materialized views for analytics.
-- Materialized views are refreshed nightly via pg_cron (Phase 9).
-- =============================================================================

-- =============================================================================
-- v_event_summary — per-event financial + attendance snapshot
-- =============================================================================
create or replace view public.v_event_summary as
select
  e.id              as event_id,
  e.slug,
  e.name,
  e.event_date,
  e.series,
  e.status,
  e.venue_id,
  coalesce(rev.revenue, 0)        as revenue,
  coalesce(exp.expenses, 0)       as expenses,
  coalesce(rev.revenue, 0) - coalesce(exp.expenses, 0) as net,
  coalesce(spn.sponsor_count, 0)  as sponsor_count,
  coalesce(spn.sponsor_cash, 0)   as sponsor_cash,
  coalesce(tkt.tickets_sold, 0)   as tickets_sold,
  coalesce(tkt.ticket_revenue, 0) as ticket_revenue,
  e.total_attendance,
  coalesce(dn.third_party_total, 0) as third_party_donations
from public.events e
left join lateral (
  select sum(amount) as revenue from public.income where event_id = e.id and deleted_at is null
) rev on true
left join lateral (
  select sum(amount) as expenses from public.expenses where event_id = e.id and deleted_at is null
) exp on true
left join lateral (
  select count(*) as sponsor_count, sum(cash_amount + in_kind_value) as sponsor_cash
  from public.sponsorships where event_id = e.id and status != 'cancelled'
) spn on true
left join lateral (
  select count(*) as tickets_sold, sum(amount_paid) as ticket_revenue
  from public.ticketing where event_id = e.id
) tkt on true
left join lateral (
  select sum(amount) as third_party_total from public.third_party_donations where event_id = e.id
) dn on true
where e.deleted_at is null;

-- =============================================================================
-- v_artist_history — per-artist booking history + earnings
-- =============================================================================
create or replace view public.v_artist_history as
select
  a.id              as artist_id,
  a.stage_name,
  count(ab.id)      as bookings_count,
  count(distinct ab.event_id) as distinct_events,
  coalesce(sum(ab.fee), 0)    as total_earned,
  coalesce(sum(ab.fee) filter (where ab.paid), 0) as total_paid,
  max(e.event_date) as last_booking_date
from public.artists a
left join public.artist_bookings ab on ab.artist_id = a.id
left join public.events e on e.id = ab.event_id and e.deleted_at is null
where a.deleted_at is null
group by a.id, a.stage_name;

-- =============================================================================
-- v_sponsor_history — per-sponsor relationship history
-- =============================================================================
create or replace view public.v_sponsor_history as
select
  s.id              as sponsor_id,
  s.name,
  count(distinct sp.event_id) as events_sponsored,
  coalesce(sum(sp.cash_amount), 0)  as total_cash,
  coalesce(sum(sp.in_kind_value), 0) as total_in_kind,
  max(e.event_date) as last_event_date,
  array_agg(distinct sp.tier) filter (where sp.tier is not null) as tiers
from public.sponsors s
left join public.sponsorships sp on sp.sponsor_id = s.id and sp.status != 'cancelled'
left join public.events e on e.id = sp.event_id and e.deleted_at is null
where s.deleted_at is null
group by s.id, s.name;

-- =============================================================================
-- v_mailing_list_health — per-segment list health
-- =============================================================================
create or replace view public.v_mailing_list_health as
select
  coalesce(seg.segment, 'all') as segment,
  count(s.id) filter (where s.status = 'subscribed') as subscribed_count,
  count(s.id) filter (where s.status = 'pending')    as pending_count,
  count(s.id) filter (where s.status = 'unsubscribed') as unsubscribed_count,
  count(s.id) filter (where s.status = 'bounced')    as bounced_count
from public.subscribers s
left join public.subscriber_segments seg on seg.subscriber_id = s.id
group by seg.segment;

-- =============================================================================
-- v_equipment_roi — per-item ROI snapshot
-- =============================================================================
create or replace view public.v_equipment_roi as
select
  ei.id             as equipment_id,
  ei.name,
  ei.category,
  ei.purchase_price,
  ei.purchase_date,
  count(eu.id)      as times_used,
  coalesce(sum(eu.revenue_attributed), 0) as total_revenue,
  case
    when ei.purchase_price is not null and ei.purchase_price > 0
    then coalesce(sum(eu.revenue_attributed), 0) / ei.purchase_price
    else null
  end as roi_ratio
from public.equipment_inventory ei
left join public.equipment_usage eu on eu.equipment_id = ei.id
where ei.deleted_at is null
group by ei.id, ei.name, ei.category, ei.purchase_price, ei.purchase_date;

-- =============================================================================
-- Materialized views — refreshed nightly via pg_cron
-- =============================================================================

create materialized view public.mv_cross_event_kpis as
select
  date_trunc('year', e.event_date)::date as year,
  e.series,
  count(*) filter (where e.status = 'completed')       as events_completed,
  coalesce(sum(rev.revenue), 0)                        as total_revenue,
  coalesce(sum(exp.expenses), 0)                       as total_expenses,
  coalesce(sum(rev.revenue), 0) - coalesce(sum(exp.expenses), 0) as net,
  coalesce(sum(e.total_attendance), 0)                 as total_attendance
from public.events e
left join lateral (
  select sum(amount) as revenue from public.income where event_id = e.id and deleted_at is null
) rev on true
left join lateral (
  select sum(amount) as expenses from public.expenses where event_id = e.id and deleted_at is null
) exp on true
where e.deleted_at is null
group by date_trunc('year', e.event_date), e.series;

create unique index idx_mv_cross_event_kpis_unique
  on public.mv_cross_event_kpis(year, series);

create materialized view public.mv_repeat_sponsors as
select
  s.id              as sponsor_id,
  s.name,
  count(distinct sp.event_id) as events_sponsored,
  sum(sp.cash_amount)         as total_cash
from public.sponsors s
join public.sponsorships sp on sp.sponsor_id = s.id and sp.status != 'cancelled'
where s.deleted_at is null
group by s.id, s.name
having count(distinct sp.event_id) >= 2;

create unique index idx_mv_repeat_sponsors_unique on public.mv_repeat_sponsors(sponsor_id);

create materialized view public.mv_top_artists as
select
  a.id              as artist_id,
  a.stage_name,
  count(distinct ab.event_id) as bookings_count,
  sum(ab.fee)       as total_earned
from public.artists a
join public.artist_bookings ab on ab.artist_id = a.id
where a.deleted_at is null
group by a.id, a.stage_name
order by count(distinct ab.event_id) desc, sum(ab.fee) desc nulls last;

create unique index idx_mv_top_artists_unique on public.mv_top_artists(artist_id);
