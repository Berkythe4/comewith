-- ============================================================
-- COME WITH — 157 cash_source on income, and a float that only counts real cash
--
-- 156 left revenue_in as "all Come With revenue", which overstates the float:
-- money only lands in the $5,000 pot if it actually arrived in PayPal or the
-- business bank. Ticket and sponsor money for an event may never have gone near
-- either. Same discipline as the outflow side — say where it moved, or it does
-- not count.
-- ============================================================
begin;

alter table public.income add column if not exists cash_source text;
alter table public.income drop constraint if exists income_cash_source_check;
alter table public.income add constraint income_cash_source_check
  check (cash_source is null or cash_source in ('paypal', 'bank', 'personal', 'other'));

comment on column public.income.cash_source is
  'Where the money actually landed. Only paypal and bank add to the business float.';

drop view if exists public.v_cash_position;
create view public.v_cash_position as
with cap as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions
   where deleted_at is null and kind = 'cash'
), inflow as (
  select coalesce(sum(amount), 0) as amt from public.income
   where deleted_at is null and ledger = 'come_with' and cash_source in ('paypal', 'bank')
), outflow as (
  select coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source in ('paypal', 'bank')
), unknown_src as (
  select count(*) as n, coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source is null
)
select cap.amt                            as capital_in,
       inflow.amt                         as revenue_in,
       outflow.amt                        as spent_from_float,
       cap.amt + inflow.amt - outflow.amt as cash_reserve,
       unknown_src.n                      as unknown_source_rows,
       unknown_src.amt                    as unknown_source_amount
  from cap, inflow, outflow, unknown_src;

revoke select on public.v_cash_position from anon;

commit;

-- DOWN: restore 156's v_cash_position; alter table public.income drop column if exists cash_source;
