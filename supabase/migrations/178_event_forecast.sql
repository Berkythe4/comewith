-- ============================================================
-- COME WITH — 178 forecast lines: money that is planned, not promised
--
-- 177 gave both ledgers three states, but all three are COMMITMENTS. `accrued`
-- already means "we have agreed to this", which is why the P&L counts it. There
-- was still nowhere to put the line before that: the DJ we intend to book, the
-- bar take we expect, the venue we are still negotiating. Those are plans, and
-- a plan is not a liability.
--
-- WHY NOT A FOURTH STATUS. Adding `planned` to expenses.status would have been
-- three lines of DDL and a permanent hazard: every view that sums expenses or
-- income WITHOUT a status filter would silently start counting speculation as
-- fact - v_pl_monthly, v_tax_year, v_event_summary, the KPI views, the 011/022/
-- 026/043/060 views. One missed filter and a forecast is in the P&L, which is
-- the exact failure LEARNINGS §33 exists to prevent. A separate table cannot
-- leak, because nothing that computes the P&L reads it.
--
-- WHY budget_lines. It already exists (026), already carries
-- (event_id, scope, category, direction, planned_amount), already has an
-- admin-only RLS policy, and its `scope` check has permitted 'event' since the
-- day it was written - the 37 rows in it today are all scope='period'. And the
-- one view that turns budget into a P&L column, v_pl_monthly_vs_budget, filters
-- `scope = 'period'`, so event-scoped lines are invisible to it by construction.
-- This migration is therefore additive to a table nothing currently writes.
--
-- REALISING A FORECAST. When the DJ is actually booked, the forecast line does
-- not get deleted - it is stamped with the id of the income/expense row it
-- became. That keeps the estimate next to the outcome, which is the only way the
-- forecast ever gets better. A realised line stops counting as forecast (it is
-- now a real row) but stays readable as history.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. What a forecast line needs that a budget line did not
-- ---------------------------------------------------------------
alter table public.budget_lines add column if not exists label text;
alter table public.budget_lines add column if not exists confidence smallint;
alter table public.budget_lines add column if not exists deleted_at timestamptz;
alter table public.budget_lines add column if not exists created_by uuid references public.profiles(id);
alter table public.budget_lines add column if not exists realized_at timestamptz;
alter table public.budget_lines add column if not exists realized_income_id uuid references public.income(id) on delete set null;
alter table public.budget_lines add column if not exists realized_expense_id uuid references public.expenses(id) on delete set null;

alter table public.budget_lines drop constraint if exists budget_lines_confidence_check;
alter table public.budget_lines add constraint budget_lines_confidence_check
  check (confidence is null or (confidence >= 0 and confidence <= 100));

-- A line becomes exactly one thing. Both set would double-count it on the way in.
alter table public.budget_lines drop constraint if exists budget_lines_realized_one_check;
alter table public.budget_lines add constraint budget_lines_realized_one_check
  check (realized_income_id is null or realized_expense_id is null);

comment on column public.budget_lines.label is
  'What the line is - "Headline DJ", "Bar minimum". The category says which P&L '
  'line it will land on; this says which actual thing it is.';
comment on column public.budget_lines.confidence is
  'Optional 0-100. Same idea as events.confidence for Blue Sky: ten $1,000 lines '
  'at 30% is $3,000 of forecast, not $10,000.';
comment on column public.budget_lines.realized_at is
  'Set when this forecast became a real income/expense row. A realised line no '
  'longer counts as forecast - it would be double-counted against the row it '
  'became - but is kept so the estimate can be compared to the outcome.';

create index if not exists idx_budget_lines_event_open
  on public.budget_lines(event_id) where deleted_at is null and realized_at is null;

-- ---------------------------------------------------------------
-- 2. The forecast, per event
-- ---------------------------------------------------------------
-- OPEN lines only. Once a line is realised the money is on the books as a real
-- row, and a forecast that keeps counting after it comes true is just an
-- overstatement with a good excuse.
create or replace view public.v_event_forecast as
select
  b.event_id,
  b.direction,
  b.category,
  count(*)                                                          as line_count,
  round(sum(b.planned_amount), 2)                                   as planned,
  round(sum(b.planned_amount * coalesce(b.confidence, 100) / 100.0), 2) as weighted,
  count(*) filter (where b.confidence is not null)                  as with_confidence
from public.budget_lines b
where b.scope = 'event' and b.event_id is not null
  and b.deleted_at is null and b.realized_at is null
group by b.event_id, b.direction, b.category;

revoke select on public.v_event_forecast from anon;

-- ---------------------------------------------------------------
-- 3. Event money carries the forecast alongside the real thing
-- ---------------------------------------------------------------
-- Appended columns only; the existing list and its order are untouched. These
-- are the ONLY place forecast money meets recorded money, and they are kept in
-- their own columns rather than folded into `revenue` / `expenses` so that no
-- existing consumer of this view can accidentally spend a plan.
create or replace view public.v_event_money as
select
  e.id as event_id, e.name, e.series, e.event_date, e.status,
  case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end as ledger,
  coalesce(t.amt, 0) as ticket_revenue,
  coalesce(s.amt, 0) as sponsor_cash,
  coalesce(d.amt, 0) as donations,
  coalesce(i.amt, 0) as other_income,
  coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) as revenue,
  coalesce(x.amt, 0) as expenses,
  coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) - coalesce(x.amt,0) as net,
  (coalesce(x.amt,0) > 0
     and coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) = 0
     and e.event_date <= current_date
     and e.revenue_confirmed_at is null)   as missing_revenue,
  (e.event_date > current_date)            as upcoming,
  (e.revenue_confirmed_at is not null)     as revenue_confirmed,
  coalesce(acc.amt, 0)                     as accrued_revenue,
  coalesce(xp.amt, 0)                      as expenses_paid,
  coalesce(xo.amt, 0)                      as expenses_owed,
  coalesce(xo.n, 0)                        as expenses_owed_count,
  coalesce(i.amt, 0) - coalesce(acc.amt, 0) as other_income_received,
  coalesce(fr.amt, 0)                      as forecast_revenue,
  coalesce(fc.amt, 0)                      as forecast_cost,
  coalesce(fr.amt, 0) - coalesce(fc.amt, 0) as forecast_net,
  coalesce(fr.wt, 0)                       as forecast_revenue_weighted,
  coalesce(fr.n, 0) + coalesce(fc.n, 0)    as forecast_line_count
from public.events e
left join lateral (select sum(amount_paid) amt from public.ticketing where event_id = e.id) t on true
left join lateral (select sum(cash_amount) amt from public.sponsorships where event_id = e.id and status <> 'cancelled') s on true
left join lateral (select sum(amount) amt from public.third_party_donations where event_id = e.id) d on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null) i on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null and status <> 'received') acc on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null) x on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null and status = 'paid') xp on true
left join lateral (select sum(amount) amt, count(*) n from public.expenses where event_id = e.id and deleted_at is null and status <> 'paid') xo on true
left join lateral (select sum(planned) amt, sum(weighted) wt, sum(line_count) n from public.v_event_forecast where event_id = e.id and direction = 'income') fr on true
left join lateral (select sum(planned) amt, sum(weighted) wt, sum(line_count) n from public.v_event_forecast where event_id = e.id and direction = 'expense') fc on true
where e.deleted_at is null;

revoke select on public.v_event_money from anon;

-- ---------------------------------------------------------------
-- 4. The old variance view learns about soft deletes
-- ---------------------------------------------------------------
-- It is revoked from anon AND authenticated (055), so nothing reads it today.
-- Fixing it now anyway: a view that would resurrect deleted rows the moment
-- someone re-granted it is a trap left lying in the schema.
create or replace view public.v_budget_variance as
select b.event_id,
       b.category,
       b.direction,
       sum(b.planned_amount) as planned,
       coalesce((select sum(e.amount) from public.expenses e
                  where e.event_id = b.event_id and e.deleted_at is null
                    and lower(e.category) = lower(b.category)), 0) as actual_expense,
       sum(b.planned_amount) -
         coalesce((select sum(e.amount) from public.expenses e
                    where e.event_id = b.event_id and e.deleted_at is null
                      and lower(e.category) = lower(b.category)), 0) as variance
  from public.budget_lines b
 where b.direction = 'expense' and b.scope = 'event' and b.deleted_at is null
 group by b.event_id, b.category, b.direction;

revoke all on public.v_budget_variance from anon, authenticated;   -- as 055 left it

commit;

-- DOWN: restore 177's v_event_money and 026's v_budget_variance;
--   drop view public.v_event_forecast;
--   alter table public.budget_lines drop column if exists label, drop column if
--   exists confidence, drop column if exists deleted_at, drop column if exists
--   created_by, drop column if exists realized_at, drop column if exists
--   realized_income_id, drop column if exists realized_expense_id;
