-- ============================================================
-- COME WITH — 153 a real P&L: revenue, direct, indirect, profit
--
-- TWO PROBLEMS WITH 149's v_pl_monthly, both found by reading it against the
-- rest of this database rather than against itself.
--
-- 1. IT WAS MISSING NEARLY ALL REVENUE. It summed `income` (1 row, $1.89) and
--    ignored ticketing ($4,005.50 over 61 rows), sponsorships ($6,225) and
--    third-party donations ($2,817.44). Migration 022 already fixed the
--    canonical revenue basis for this business and 149 simply did not use it:
--      gross_revenue = ticket_revenue + other_income + donations + sponsor_cash
--    So the P&L showed a business that only ever spent money. ~$13,048 missing.
--
-- 2. IT WAS A FLAT LIST. Every category in one column, so there was no gross
--    margin and no way to see whether events pay for themselves before overhead.
--    A P&L has a shape: revenue, the direct cost of earning it, then the cost of
--    existing at all.
--
--      Revenue
--    - Direct    (cost tied to a specific event)
--    = Gross profit
--    - Indirect  (overhead - event_na, or simply not event-linked)
--    = Net profit
--
-- v_pl_monthly is REPLACED with the same column names plus `section`, so the
-- existing callers keep working and gain the structure.
--
-- Also adds v_cash_position. Cash is not the same question as capital: capital
-- is what Keith put in, cash is what the business can actually spend.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. The P&L, sectioned, with every revenue stream this business has
-- ---------------------------------------------------------------
create or replace view public.v_pl_monthly as
-- REVENUE, by stream. Each keeps its own identity so "where does money come
-- from" is answerable, not just "how much".
with rev as (
  select to_char(t.purchased_at, 'YYYY-MM') as period, 'Ticket sales'::text as category,
         sum(t.amount_paid) as revenue, 0::numeric as cost, count(*) as n
    from public.ticketing t where t.amount_paid is not null group by 1
  union all
  select to_char(coalesce(e.event_date, s.created_at::date), 'YYYY-MM'), 'Sponsorship',
         sum(s.cash_amount), 0, count(*)
    from public.sponsorships s
    left join public.events e on e.id = s.event_id
   where s.status <> 'cancelled' and s.cash_amount is not null group by 1
  union all
  select to_char(coalesce(d.date, e.event_date, d.created_at::date), 'YYYY-MM'), 'Donations',
         sum(d.amount), 0, count(*)
    from public.third_party_donations d
    left join public.events e on e.id = d.event_id
   where d.amount is not null group by 1
  union all
  select to_char(i.date, 'YYYY-MM'), coalesce(nullif(i.category, ''), 'Other income'),
         sum(i.amount), 0, count(*)
    from public.income i where i.deleted_at is null group by 1, 2
),
-- COST, split on whether it belongs to a specific event. `event_na` is the
-- explicit overhead marker added in 050; anything else without an event is
-- treated as indirect too, since it demonstrably is not attached to one.
cost as (
  select to_char(date, 'YYYY-MM') as period, coalesce(nullif(category,''), 'Uncategorised') as category,
         0::numeric as revenue, sum(amount) as cost, count(*) as n,
         (event_id is not null) as is_direct
    from public.expenses where deleted_at is null
   group by 1, 2, 6
)
-- Column ORDER must match 149's exactly and `section` can only be APPENDED:
-- create-or-replace may add trailing columns but never rename or reorder them,
-- and v_pl_monthly_vs_budget already reads this view.
select period, category,
       'revenue'::text as bucket,          -- 149 callers read `bucket`
       'business'::text as funded_by,
       sum(revenue) as revenue, sum(cost) as cost, sum(revenue) - sum(cost) as net,
       sum(n) as line_count,
       'revenue'::text as section
  from rev group by 1, 2
union all
select period, category,
       case when is_direct then 'event' else 'overhead' end,
       'business',
       sum(revenue), sum(cost), sum(revenue) - sum(cost), sum(n),
       case when is_direct then 'direct' else 'indirect' end
  from cost group by 1, 2, is_direct;

revoke select on public.v_pl_monthly from anon;

-- ---------------------------------------------------------------
-- 2. The three profit lines, per month
-- ---------------------------------------------------------------
create or replace view public.v_pl_summary as
select period,
       round(sum(revenue) filter (where section = 'revenue'), 2)            as revenue,
       round(sum(cost)    filter (where section = 'direct'), 2)             as direct_cost,
       round(sum(cost)    filter (where section = 'indirect'), 2)           as indirect_cost,
       round(sum(revenue) filter (where section = 'revenue')
           - sum(cost)    filter (where section = 'direct'), 2)             as gross_profit,
       round(sum(revenue) filter (where section = 'revenue')
           - sum(cost)    filter (where section in ('direct','indirect')), 2) as net_profit
  from public.v_pl_monthly
 group by period;

revoke select on public.v_pl_summary from anon;

-- ---------------------------------------------------------------
-- 3. Cash, which is not capital
-- ---------------------------------------------------------------
-- Capital = what Keith put in. Cash = what the business can spend. Expenses he
-- paid personally never touched business cash, so they are excluded here even
-- though they are very much real costs in the P&L above.
--
-- CAVEAT, and it matters: funded_by defaulted to 'business' for every row that
-- pre-dates the Jennifer push. 67 site-entered expenses have never been reviewed,
-- so some business-paid spend here was probably Keith's. A negative cash figure
-- is the symptom - use the Expenses tab's funding filter to correct it.
create or replace view public.v_cash_position as
with cap as (
  select coalesce(sum(amount), 0) as amt
    from public.capital_contributions where deleted_at is null and kind = 'cash'
), inflow as (
  select coalesce(sum(revenue), 0) as amt
    from public.v_pl_monthly where section = 'revenue'
), outflow as (
  select coalesce(sum(amount), 0) as amt
    from public.expenses where deleted_at is null and funded_by = 'business'
), unreviewed as (
  select count(*) as n, coalesce(sum(amount), 0) as amt
    from public.expenses
   where deleted_at is null and funded_by = 'business' and external_ref is null
)
select cap.amt                                   as capital_in,
       inflow.amt                                as revenue_in,
       outflow.amt                               as spent_by_business,
       cap.amt + inflow.amt - outflow.amt        as cash_reserve,
       unreviewed.n                              as unreviewed_funding_rows,
       unreviewed.amt                            as unreviewed_funding_amount
  from cap, inflow, outflow, unreviewed;

revoke select on public.v_cash_position from anon;

commit;

-- DOWN: restore 149's v_pl_monthly definition, then
--   drop view if exists public.v_cash_position;
--   drop view if exists public.v_pl_summary;
