-- ============================================================
-- COME WITH — 177 payables: the expense side of the accrual
--
-- 161 gave INCOME three states (accrued -> invoiced -> received) because most of
-- what Come With earns is agreed long before it is paid. Costs work the same way
-- and had no way to say so: a DJ booked in August and paid in October was either
-- invisible or recorded as though the money had already left. The first
-- understates cost, the second understates cash. Both are wrong at once.
--
-- expenses.status carries it, mirroring income:
--   accrued   committed, no bill yet   (we have agreed to pay them)
--   invoiced  bill received, unpaid    (they have asked)
--   paid      money left               (today's only state)
--
-- WHAT COUNTS WHERE — the whole point of the split:
--   P&L      counts all three. A cost is incurred when the obligation is, not
--            when the cheque clears. v_pl_monthly is unchanged for that reason.
--   CASH     counts `paid` only. v_cash_position's outflow is now filtered, or
--            recording a payable would silently drain the float.
--   1099     counts `paid` only, in the year of PAYMENT. A 1099 is cash-basis:
--            a fee accrued in December and paid in January is next year's form.
--            Hence coalesce(settled_at::date, date) as the tax year, here and
--            in v_tax_year.
--   CAPITAL  counts `paid` only. Keith cannot have personally carried a bill
--            that nobody has paid yet.
--
-- EVERY EXISTING ROW BACKFILLS TO 'paid' with settled_at null, so the coalesce
-- falls back to `date` and every number this database reports today is
-- byte-identical after this migration. Nothing moves until someone records a
-- payable on purpose.
--
-- Additive: no column dropped, every new/replaced view anon-revoked (E1).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Accrual state on expenses
-- ---------------------------------------------------------------
alter table public.expenses add column if not exists status text not null default 'paid';
alter table public.expenses drop constraint if exists expenses_status_check;
alter table public.expenses add constraint expenses_status_check
  check (status in ('accrued', 'invoiced', 'paid'));

alter table public.expenses add column if not exists expected_amount numeric(10,2);
alter table public.expenses add column if not exists settled_at timestamptz;
alter table public.expenses add column if not exists due_date date;

comment on column public.expenses.status is
  'accrued = committed, no bill yet. invoiced = billed, unpaid. paid = money '
  'left. The P&L counts all three (a cost is incurred when the obligation is); '
  'only paid counts as cash, as capital, or toward a 1099.';
comment on column public.expenses.expected_amount is
  'What was agreed, kept when the final settles differently, so the variance '
  'survives being corrected.';
comment on column public.expenses.settled_at is
  'When the money actually left. Null while unpaid. Drives the CASH-BASIS year '
  'for 1099s and the tax summary, which is not always the year it was incurred.';
comment on column public.expenses.due_date is
  'When this is owed. Null means no date agreed - not "not overdue".';

create index if not exists idx_expenses_status on public.expenses(status);
create index if not exists idx_expenses_due_date on public.expenses(due_date)
  where status <> 'paid';

-- ---------------------------------------------------------------
-- 2. What Come With owes — the mirror of v_receivables
-- ---------------------------------------------------------------
create or replace view public.v_payables as
select x.id, x.date, x.due_date, x.amount, x.expected_amount, x.status,
       x.category, x.vendor, x.ledger, x.description,
       coalesce(a.display_name, x.vendor)              as payee,
       x.vendor_actor_id,
       e.name                                          as event_name,
       e.event_date,
       x.event_id,
       (current_date - x.date)                         as days_since_incurred,
       case when x.due_date is null then null
            else (current_date - x.due_date) end       as days_overdue,
       (x.due_date is not null and x.due_date < current_date) as overdue
  from public.expenses x
  left join public.actors a on a.id = x.vendor_actor_id
  left join public.events e on e.id = x.event_id and e.deleted_at is null
 where x.deleted_at is null and x.status in ('accrued', 'invoiced');

revoke select on public.v_payables from anon;

-- ---------------------------------------------------------------
-- 3. Cash counts only money that actually moved — both directions
-- ---------------------------------------------------------------
drop view if exists public.v_cash_position;
create view public.v_cash_position as
with cap as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions
   where deleted_at is null and kind = 'cash'
), inflow as (
  select coalesce(sum(amount), 0) as amt from public.income
   where deleted_at is null and ledger = 'come_with'
     and status = 'received' and cash_source in ('paypal', 'bank')
), outflow as (
  -- status filter added by 177. Without it a payable recorded today reads as
  -- money already gone, and the float understates itself by whatever is owed.
  select coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with'
     and status = 'paid' and cash_source in ('paypal', 'bank')
), unknown_src as (
  -- An unpaid bill has no cash source yet BY DEFINITION. Nagging about it would
  -- make the "no source" queue grow every time someone plans ahead properly.
  select count(*) as n, coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with'
     and status = 'paid' and cash_source is null
), owed as (
  select coalesce(sum(amount), 0) as amt, count(*) as n from public.v_receivables
   where ledger = 'come_with'
), owing as (
  select coalesce(sum(amount), 0) as amt, count(*) as n from public.v_payables
   where ledger = 'come_with'
)
select cap.amt                            as capital_in,
       inflow.amt                         as revenue_in,
       outflow.amt                        as spent_from_float,
       cap.amt + inflow.amt - outflow.amt as cash_reserve,
       unknown_src.n                      as unknown_source_rows,
       unknown_src.amt                    as unknown_source_amount,
       owed.amt                           as owed_to_us,
       owed.n                             as owed_count,
       owing.amt                          as owed_by_us,
       owing.n                            as owing_count,
       cap.amt + inflow.amt - outflow.amt - owing.amt + owed.amt as cash_after_commitments
  from cap, inflow, outflow, unknown_src, owed, owing;

revoke select on public.v_cash_position from anon;

-- ---------------------------------------------------------------
-- 4. Event money splits cost into paid and owed
-- ---------------------------------------------------------------
-- Columns are APPENDED; the existing list and its order are untouched so every
-- consumer of this view keeps reading what it read yesterday. `expenses` stays
-- the total incurred (the P&L number) — an event's cost does not drop because
-- the invoice is still sitting in a drawer.
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
  coalesce(i.amt, 0) - coalesce(acc.amt, 0) as other_income_received
from public.events e
left join lateral (select sum(amount_paid) amt from public.ticketing where event_id = e.id) t on true
left join lateral (select sum(cash_amount) amt from public.sponsorships where event_id = e.id and status <> 'cancelled') s on true
left join lateral (select sum(amount) amt from public.third_party_donations where event_id = e.id) d on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null) i on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null and status <> 'received') acc on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null) x on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null and status = 'paid') xp on true
left join lateral (select sum(amount) amt, count(*) n from public.expenses where event_id = e.id and deleted_at is null and status <> 'paid') xo on true
where e.deleted_at is null;

revoke select on public.v_event_money from anon;

-- ---------------------------------------------------------------
-- 5. A 1099 reports what was PAID, in the year it was paid
-- ---------------------------------------------------------------
drop view if exists public.v_contractor_1099;

create view public.v_contractor_1099 as
select
  coalesce(a.display_name, x.vendor)            as payee,
  x.vendor_actor_id                             as actor_id,
  extract(year from coalesce(x.settled_at::date, x.date))::int as tax_year,
  count(*)                                      as payments,
  round(sum(x.amount), 2)                       as total_paid,
  min(coalesce(x.settled_at::date, x.date))     as first_payment,
  max(coalesce(x.settled_at::date, x.date))     as last_payment,
  string_agg(distinct x.category, ', ' order by x.category) as categories,
  (sum(x.amount) >= 600)                        as over_threshold,
  round(greatest(600 - sum(x.amount), 0), 2)    as headroom,
  case when x.vendor_actor_id is null then 'no vendor'
       else coalesce(a.tax_1099_status, 'undecided') end as status,
  a.tax_1099_note                               as note,
  (sum(x.amount) >= 600
   and (x.vendor_actor_id is null or a.tax_1099_status is null)) as needs_review
from public.expenses x
left join public.actors a on a.id = x.vendor_actor_id
where x.deleted_at is null
  and x.ledger = 'come_with'
  and x.status = 'paid'
group by 1, 2, 3, a.tax_1099_status, a.tax_1099_note;

revoke select on public.v_contractor_1099 from anon;

-- ---------------------------------------------------------------
-- 6. Capital: an unpaid bill was not carried by anybody
-- ---------------------------------------------------------------
create or replace view public.v_capital as
with contrib as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions where deleted_at is null
), personal as (
  select coalesce(sum(amount), 0) as amt,
         coalesce(sum(amount) filter (where reimbursed_at is not null), 0) as repaid
    from public.expenses
   where deleted_at is null and funded_by = 'owner' and status = 'paid'
)
select
  contrib.amt                                as contributed,
  personal.amt                               as personally_paid,
  personal.repaid                            as reimbursed,
  contrib.amt + personal.amt                 as invested_gross,
  contrib.amt + personal.amt - personal.repaid as invested_net,
  personal.amt - personal.repaid             as outstanding_reimbursable
from contrib, personal;

revoke select on public.v_capital from anon;

-- ---------------------------------------------------------------
-- 7. Tax year, on the cash basis it is actually filed on
-- ---------------------------------------------------------------
create or replace view public.v_tax_year as
select
  extract(year from coalesce(settled_at::date, date))::int    as tax_year,
  category,
  -- No coalesce anywhere below: 151 returned null for "this category has no such
  -- rows" and the dashboard reads that as blank. Wrapping it in coalesce would
  -- turn "none of this category is non-deductible" into a measured $0.00, which
  -- is a different claim (LEARNINGS §23). The ONLY change here is the filter.
  count(*) filter (where status = 'paid')                     as line_count,
  round(sum(amount) filter (where status = 'paid'), 2)                     as total,
  round(sum(amount) filter (where status = 'paid' and deductible), 2)      as deductible_total,
  round(sum(amount) filter (where status = 'paid' and not deductible), 2)  as non_deductible_total,
  round(sum(amount) filter (where status = 'paid' and funded_by = 'owner'), 2) as paid_personally,
  count(*) filter (where status = 'paid' and receipt_path is null)  as missing_receipts,
  -- Committed but not yet paid. Not deductible on a cash basis and deliberately
  -- kept out of every total above — shown so it is not invisible either.
  round(coalesce(sum(amount) filter (where status <> 'paid'), 0), 2) as committed_unpaid
from public.expenses
where deleted_at is null
group by 1, 2;

revoke select on public.v_tax_year from anon;

-- ---------------------------------------------------------------
-- 8. Health: an overdue bill is a thing that needs a human
-- ---------------------------------------------------------------
create or replace view public.v_expense_health as
select
  e.id, e.date, e.amount, e.category, e.vendor, e.event_id, e.event_na,
  e.funded_by, e.verified_at, e.receipt_path,
  (e.category is null)                                    as needs_category,
  (e.event_id is null and not e.event_na)                 as needs_event,
  (e.verified_at is null)                                 as needs_review,
  (e.receipt_path is null and e.amount >= 100)            as needs_receipt,
  exists (select 1 from public.expenses d
           where d.id <> e.id and d.deleted_at is null
             and d.date = e.date and d.amount = e.amount)  as possible_duplicate,
  e.status                                                as status,
  e.due_date                                              as due_date,
  (e.status <> 'paid')                                    as unpaid,
  (e.status <> 'paid' and e.due_date is not null and e.due_date < current_date) as overdue
from public.expenses e
where e.deleted_at is null;

revoke select on public.v_expense_health from anon;

commit;

-- DOWN: restore 171's v_contractor_1099, 161's v_cash_position + v_event_money,
--   151's v_capital / v_tax_year / v_expense_health; drop view public.v_payables;
--   alter table public.expenses drop column if exists status, drop column if
--   exists expected_amount, drop column if exists settled_at, drop column if
--   exists due_date;
