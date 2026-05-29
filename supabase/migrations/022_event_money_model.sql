-- ============================================================
-- COME WITH — 022 CANONICAL EVENT MONEY MODEL
--
-- One consistent revenue basis for every event; workstreams differ only in
-- framing. Locked definitions:
--
--   ticket_revenue  = Σ ticketing.amount_paid
--   other_income    = Σ income.amount        (event-linked, deleted filtered; NON-ticket)
--   donations       = Σ third_party_donations.amount
--   sponsor_cash    = Σ sponsorships.cash_amount     (excl. cancelled)
--   sponsor_in_kind = Σ sponsorships.in_kind_value   (excl. cancelled)
--   total_expenses  = Σ expenses.amount       (event-linked, deleted filtered)
--   gross_revenue   = ticket_revenue + other_income + donations + sponsor_cash
--
--   EVERY event:        net_pl       = gross_revenue − total_expenses
--   Dance Infusion adds: total_raised = gross_revenue + sponsor_in_kind
--                        cost_to_raise = total_expenses / total_raised
--
-- Change: split sponsor_cash vs in_kind in v_event_summary (sponsor_cash was
-- cash+in_kind); rebuild the rollup with gross_revenue / net_pl / total_raised;
-- repoint v_kpi_parties (net_pl from rollup) and v_kpi_dance_infusion (net_pl +
-- total_raised + cost_to_raise). Re-assert anon revoke on every touched view
-- (015/019 regression guard — these are financial views).
-- ============================================================
begin;

-- 1. v_event_summary: sponsor_cash becomes cash-only; add sponsor_in_kind (appended).
--    'net' (income − expenses) and all other columns are unchanged.
create or replace view public.v_event_summary as
select
  e.id as event_id, e.slug, e.name, e.event_date, e.series, e.status, e.venue_id,
  coalesce(rev.revenue, 0)  as revenue,
  coalesce(exp.expenses, 0) as expenses,
  coalesce(rev.revenue, 0) - coalesce(exp.expenses, 0) as net,
  coalesce(spn.sponsor_count, 0) as sponsor_count,
  coalesce(spn.sponsor_cash, 0)  as sponsor_cash,
  coalesce(tkt.tickets_sold, 0)  as tickets_sold,
  coalesce(tkt.ticket_revenue, 0) as ticket_revenue,
  e.total_attendance,
  coalesce(dn.third_party_total, 0) as third_party_donations,
  coalesce(spn.sponsor_in_kind, 0)  as sponsor_in_kind
from public.events e
left join lateral (select sum(amount) as revenue  from public.income   where event_id=e.id and deleted_at is null) rev on true
left join lateral (select sum(amount) as expenses from public.expenses where event_id=e.id and deleted_at is null) exp on true
left join lateral (select count(*) as sponsor_count,
                          sum(cash_amount)  as sponsor_cash,
                          sum(in_kind_value) as sponsor_in_kind
                   from public.sponsorships where event_id=e.id and status <> 'cancelled') spn on true
left join lateral (select sum(coalesce(quantity,1)) as tickets_sold, sum(amount_paid) as ticket_revenue
                   from public.ticketing where event_id=e.id) tkt on true
left join lateral (select sum(amount) as third_party_total from public.third_party_donations where event_id=e.id) dn on true
where e.deleted_at is null;

revoke select on public.v_event_summary from anon;

-- 2. Rebuild the KPI views (drop leaves first so the rollup can be recolumned).
drop view if exists public.v_kpi_parties;
drop view if exists public.v_kpi_dance_infusion;
drop view if exists public.v_kpi_event_financials;

create view public.v_kpi_event_financials as
select
  s.event_id, s.name, s.series, s.event_date, e.capacity, s.total_attendance,
  s.tickets_sold, s.ticket_revenue,
  s.revenue               as other_income,
  s.expenses              as total_expenses,
  s.third_party_donations as donations,
  s.sponsor_cash          as sponsor_cash,
  s.sponsor_in_kind       as sponsor_in_kind,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash)              as gross_revenue,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash) - s.expenses as net_pl,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash + s.sponsor_in_kind) as total_raised
from public.v_event_summary s
join public.events e on e.id = s.event_id;

revoke select on public.v_kpi_event_financials from anon;

-- Come With Parties — sell-through + net P&L
create view public.v_kpi_parties as
select
  event_id, name, event_date, capacity, tickets_sold,
  case when capacity > 0 then round(tickets_sold::numeric / capacity * 100, 1) end as sell_through_pct,
  net_pl
from public.v_kpi_event_financials
where series = 'Come With Parties';

revoke select on public.v_kpi_parties from anon;

-- Dance Infusion — net P&L + total raised (incl in-kind) + cost to raise $1
create view public.v_kpi_dance_infusion as
select
  event_id, name, event_date, total_attendance,
  net_pl,
  total_raised,
  case when total_raised > 0 then round(total_expenses / nullif(total_raised, 0), 2) end as cost_to_raise_per_dollar
from public.v_kpi_event_financials
where series = 'Dance Infusion';

revoke select on public.v_kpi_dance_infusion from anon;

commit;
