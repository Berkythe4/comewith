-- ============================================================
-- COME WITH — 149 FP&A / P&L
--
-- Makes this repo the home of the Come With P&L. Two things were missing:
--
--   1. A PERIOD view. Everything financial here is per-EVENT (v_event_summary,
--      v_kpi_event_financials). But 112 of 133 expenses are overhead
--      (event_na = true) and belong to no event at all, so they never appear in
--      any rollup. A month-by-month P&L is the missing lens.
--
--   2. BUDGETS that aren't event-scoped. budget_lines exists and is EMPTY (0
--      rows); its scope check allows 'event' | 'event_type' | 'overall' but has
--      no period column, so "Software costs $230/mo" cannot be expressed.
--
-- Plus the plumbing for the Jennifer push: a stable external_ref so the same
-- charge can never land twice, and funded_by so the business can tell what Keith
-- paid for personally (his capital contribution) from what it paid itself.
--
-- WHY external_ref MATTERS — measured, not assumed. Jennifer holds 180 Come With
-- rows ($14,554.13); this database holds 133 expenses ($26,976.42). Exactly 66 of
-- them ($4,874.34) are the SAME charge recorded in both places. Pushing blind
-- would have created 66 duplicates. The unique index below makes that impossible
-- rather than merely unlikely.
--
-- Additive only: no column is dropped, no existing view is redefined, and every
-- new financial view is anon-revoked (E1 discipline / the 016-017 regression).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Idempotency + funding source on the two money ledgers
-- ---------------------------------------------------------------
alter table public.expenses add column if not exists external_ref text;
alter table public.income   add column if not exists external_ref text;

-- Partial unique: only rows that carry a ref are constrained, so the 133 rows
-- already here (external_ref null) are untouched and stay insertable by hand.
create unique index if not exists uq_expenses_external_ref
  on public.expenses(external_ref) where external_ref is not null;
create unique index if not exists uq_income_external_ref
  on public.income(external_ref) where external_ref is not null;

comment on column public.expenses.external_ref is
  'Stable id from the sending system (Jennifer). Makes the push idempotent: '
  're-sending the same charge updates rather than duplicates. Null for rows '
  'entered by hand here.';

-- Who actually paid. 'owner' = Keith carried it on his personal card, which is a
-- capital contribution to the business, not business cash going out. The
-- distinction drives what Come With owes him back.
alter table public.expenses add column if not exists funded_by text
  not null default 'business';
alter table public.expenses
  drop constraint if exists expenses_funded_by_check;
alter table public.expenses add constraint expenses_funded_by_check
  check (funded_by in ('business', 'owner'));

-- ---------------------------------------------------------------
-- 2. Period budgets
-- ---------------------------------------------------------------
alter table public.budget_lines add column if not exists period text;   -- 'YYYY-MM'
alter table public.budget_lines
  drop constraint if exists budget_lines_period_format_check;
alter table public.budget_lines add constraint budget_lines_period_format_check
  check (period is null or period ~ '^[0-9]{4}-(0[1-9]|1[0-2])$');

-- Widen the scope vocabulary to include a period-scoped line. The existing
-- constraint is replaced, not dropped-and-forgotten, so 'event' still validates.
alter table public.budget_lines
  drop constraint if exists budget_lines_scope_check;
alter table public.budget_lines add constraint budget_lines_scope_check
  check (scope in ('event', 'event_type', 'overall', 'period'));

create index if not exists idx_budget_lines_period on public.budget_lines(period);

-- ---------------------------------------------------------------
-- 3. The P&L itself — one row per month per category, actual vs budget
-- ---------------------------------------------------------------
-- Revenue and cost are kept as separate signed columns rather than one signed
-- amount: expenses.amount is stored POSITIVE here (a cost of 42.00, not -42.00),
-- and flipping that convention would silently rewrite every existing view.
create or replace view public.v_pl_monthly as
with months as (
  select to_char(date, 'YYYY-MM') as period, category,
         case when event_na then 'overhead' else 'event' end as bucket,
         funded_by,
         sum(amount) as cost,
         0::numeric   as revenue,
         count(*)     as n
    from public.expenses
   where deleted_at is null
   group by 1, 2, 3, 4
  union all
  select to_char(date, 'YYYY-MM'), coalesce(category, 'Income'),
         case when event_id is null then 'overhead' else 'event' end,
         'business',
         0::numeric,
         sum(amount),
         count(*)
    from public.income
   where deleted_at is null
   group by 1, 2, 3, 4
)
select period, category, bucket, funded_by,
       sum(revenue)              as revenue,
       sum(cost)                 as cost,
       sum(revenue) - sum(cost)  as net,
       sum(n)                    as line_count
  from months
 group by period, category, bucket, funded_by;

revoke select on public.v_pl_monthly from anon;   -- financial view, E1

-- Budget vs actual, by month and category. Left join from a full period x
-- category key set so a budgeted line with no spend still shows (that is the
-- variance you most want to see) and unbudgeted spend still shows too.
create or replace view public.v_pl_monthly_vs_budget as
with actual as (
  select period, category, sum(cost) as cost, sum(revenue) as revenue
    from public.v_pl_monthly group by period, category
), planned as (
  select period, category,
         sum(case when direction = 'expense' then planned_amount else 0 end) as planned_cost,
         sum(case when direction = 'income'  then planned_amount else 0 end) as planned_revenue
    from public.budget_lines
   where scope = 'period' and period is not null
   group by period, category
), keys as (
  select period, category from actual
  union
  select period, category from planned
)
select k.period, k.category,
       coalesce(a.revenue, 0)         as revenue,
       coalesce(a.cost, 0)            as cost,
       coalesce(p.planned_revenue, 0) as planned_revenue,
       coalesce(p.planned_cost, 0)    as planned_cost,
       coalesce(a.cost, 0) - coalesce(p.planned_cost, 0)       as cost_variance,
       coalesce(a.revenue, 0) - coalesce(p.planned_revenue, 0) as revenue_variance
  from keys k
  left join actual  a using (period, category)
  left join planned p using (period, category);

revoke select on public.v_pl_monthly_vs_budget from anon;

-- Owner-funded spend = what Come With owes Keith back. This is the number the
-- planner shows as "invested in Come With"; having it here means the two systems
-- can be reconciled instead of argued about.
create or replace view public.v_owner_funded as
select to_char(date, 'YYYY-MM') as period,
       sum(amount) as owner_funded,
       count(*)    as line_count
  from public.expenses
 where deleted_at is null and funded_by = 'owner'
 group by 1;

revoke select on public.v_owner_funded from anon;

-- ---------------------------------------------------------------
-- 4. Register the dashboard module
-- ---------------------------------------------------------------
insert into public.module_registry (key, label, nav_group, sort_order, built, master_only)
values ('pl', 'P&L', 'Finance', 15, true, true)
on conflict (key) do update
  set label = excluded.label, nav_group = excluded.nav_group, built = excluded.built;

commit;

-- DOWN:
--   drop view if exists public.v_owner_funded;
--   drop view if exists public.v_pl_monthly_vs_budget;
--   drop view if exists public.v_pl_monthly;
--   drop index if exists public.uq_expenses_external_ref;
--   drop index if exists public.uq_income_external_ref;
--   drop index if exists public.idx_budget_lines_period;
--   alter table public.expenses drop column if exists external_ref, drop column if exists funded_by;
--   alter table public.income   drop column if exists external_ref;
--   alter table public.budget_lines drop column if exists period;
--   delete from public.module_registry where key = 'pl';
