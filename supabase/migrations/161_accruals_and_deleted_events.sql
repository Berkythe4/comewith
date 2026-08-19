-- ============================================================
-- COME WITH — 161 revenue accruals, confirmed-zero events, and a real bug
--
-- BUG FIRST. v_pl_monthly reported $550 of Come With ticket sales that do not
-- exist. They come from two events named 'Test' and 'ZZ delete me' which are
-- SOFT-DELETED — but the ticketing CTE joined events without checking
-- deleted_at, so their rows kept contributing. Every other view in this database
-- filters deleted_at; 149 did not, and inherited a $550 phantom.
--
-- With that gone, Come With's own recorded revenue is $0. Not $550. Which makes
-- the rest of this migration the point rather than a nicety.
--
-- ACCRUALS. A gig is agreed in June, played in July, paid in August. Until now
-- the only recordable state was "money in hand", so agreed-but-unpaid work was
-- invisible — which is most of what Come With does. income.status carries it:
--   accrued   agreed, not invoiced
--   invoiced  billed, not paid
--   received  money arrived
-- The P&L counts all three, because revenue is earned when the work is done.
-- CASH counts only `received`, because that is what cash means. expected_amount
-- keeps the agreed figure when the final differs, so the variance survives.
--
-- CONFIRMED-ZERO EVENTS. "No revenue recorded" and "genuinely earned nothing"
-- look identical in the data and mean opposite things. An event cannot be marked
-- verified through its income rows when it HAS no income rows, so the flag lives
-- on the event.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Accrual state on income
-- ---------------------------------------------------------------
alter table public.income add column if not exists status text not null default 'received';
alter table public.income drop constraint if exists income_status_check;
alter table public.income add constraint income_status_check
  check (status in ('accrued', 'invoiced', 'received'));

alter table public.income add column if not exists expected_amount numeric(10,2);
alter table public.income add column if not exists settled_at timestamptz;

comment on column public.income.status is
  'accrued = agreed, not invoiced. invoiced = billed, not paid. received = money '
  'arrived. The P&L counts all three (revenue is earned when the work is done); '
  'only received counts as cash.';
comment on column public.income.expected_amount is
  'What was agreed, kept when the final settles differently, so the variance is '
  'not lost the moment it is corrected.';

create index if not exists idx_income_status on public.income(status);

-- ---------------------------------------------------------------
-- 2. An event can be confirmed to have earned nothing
-- ---------------------------------------------------------------
alter table public.events add column if not exists revenue_confirmed_at timestamptz;
alter table public.events add column if not exists revenue_confirmed_by uuid references public.profiles(id);

comment on column public.events.revenue_confirmed_at is
  'Set when someone has confirmed this event''s revenue is complete — including '
  'confirming it earned nothing. Distinguishes "checked, it was free" from '
  '"nobody has looked yet", which are identical in the data and opposite in meaning.';

-- ---------------------------------------------------------------
-- 3. The fix: deleted events must not contribute
-- ---------------------------------------------------------------
create or replace view public.v_pl_monthly as
with rev as (
  select to_char(coalesce(t.purchased_at::date, e.event_date, t.created_at::date), 'YYYY-MM') as period,
         'Ticket sales'::text as category, sum(t.amount_paid) as revenue, 0::numeric as cost, count(*) as n,
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end as ledger
    from public.ticketing t
    join public.events e on e.id = t.event_id and e.deleted_at is null   -- <- the fix
   where t.amount_paid is not null group by 1, 6
  union all
  select to_char(coalesce(e.event_date, s.created_at::date), 'YYYY-MM'), 'Sponsorship',
         sum(s.cash_amount), 0, count(*),
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
    from public.sponsorships s
    join public.events e on e.id = s.event_id and e.deleted_at is null
   where s.status <> 'cancelled' and s.cash_amount is not null group by 1, 6
  union all
  select to_char(coalesce(d.date, e.event_date, d.created_at::date), 'YYYY-MM'), 'Donations',
         sum(d.amount), 0, count(*),
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
    from public.third_party_donations d
    join public.events e on e.id = d.event_id and e.deleted_at is null
   where d.amount is not null group by 1, 6
  union all
  select to_char(i.date, 'YYYY-MM'), coalesce(nullif(i.category, ''), 'Other income'),
         sum(i.amount), 0, count(*), i.ledger
    from public.income i
    left join public.events e on e.id = i.event_id
   where i.deleted_at is null and (i.event_id is null or e.deleted_at is null)
   group by 1, 2, 6
),
cost as (
  select to_char(x.date, 'YYYY-MM') as period, coalesce(nullif(x.category,''), 'Uncategorised') as category,
         0::numeric as revenue, sum(x.amount) as cost, count(*) as n,
         (x.event_id is not null) as is_direct, x.ledger
    from public.expenses x
    left join public.events e on e.id = x.event_id
   where x.deleted_at is null and (x.event_id is null or e.deleted_at is null)
   group by 1, 2, 6, 7
)
select period, category,
       'revenue'::text as bucket, 'business'::text as funded_by,
       sum(revenue) as revenue, sum(cost) as cost, sum(revenue) - sum(cost) as net,
       sum(n) as line_count, 'revenue'::text as section, ledger
  from rev group by 1, 2, ledger
union all
select period, category,
       case when is_direct then 'event' else 'overhead' end, 'business',
       sum(revenue), sum(cost), sum(revenue) - sum(cost), sum(n),
       case when is_direct then 'direct' else 'indirect' end, ledger
  from cost group by 1, 2, is_direct, ledger;

revoke select on public.v_pl_monthly from anon;

-- ---------------------------------------------------------------
-- 4. What is owed to Come With
-- ---------------------------------------------------------------
create or replace view public.v_receivables as
select i.id, i.date, i.amount, i.expected_amount, i.status, i.category, i.ledger,
       i.description, e.name as event_name, e.event_date,
       (current_date - i.date) as days_outstanding
  from public.income i
  left join public.events e on e.id = i.event_id
 where i.deleted_at is null and i.status in ('accrued', 'invoiced');

revoke select on public.v_receivables from anon;

-- ---------------------------------------------------------------
-- 5. Cash counts only money that arrived
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
  select coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source in ('paypal', 'bank')
), unknown_src as (
  select count(*) as n, coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source is null
), owed as (
  select coalesce(sum(amount), 0) as amt, count(*) as n from public.v_receivables
   where ledger = 'come_with'
)
select cap.amt                            as capital_in,
       inflow.amt                         as revenue_in,
       outflow.amt                        as spent_from_float,
       cap.amt + inflow.amt - outflow.amt as cash_reserve,
       unknown_src.n                      as unknown_source_rows,
       unknown_src.amt                    as unknown_source_amount,
       owed.amt                           as owed_to_us,
       owed.n                             as owed_count
  from cap, inflow, outflow, unknown_src, owed;

revoke select on public.v_cash_position from anon;

-- ---------------------------------------------------------------
-- 6. Event money respects the confirmation flag
-- ---------------------------------------------------------------
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
  -- Spent money, earned none, in the past, AND nobody has confirmed that is correct.
  (coalesce(x.amt,0) > 0
     and coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) = 0
     and e.event_date <= current_date
     and e.revenue_confirmed_at is null)   as missing_revenue,
  (e.event_date > current_date)            as upcoming,
  (e.revenue_confirmed_at is not null)     as revenue_confirmed,
  coalesce(acc.amt, 0)                     as accrued_revenue
from public.events e
left join lateral (select sum(amount_paid) amt from public.ticketing where event_id = e.id) t on true
left join lateral (select sum(cash_amount) amt from public.sponsorships where event_id = e.id and status <> 'cancelled') s on true
left join lateral (select sum(amount) amt from public.third_party_donations where event_id = e.id) d on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null) i on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null and status <> 'received') acc on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null) x on true
where e.deleted_at is null;

revoke select on public.v_event_money from anon;

commit;

-- DOWN: restore 160's views; alter table public.income drop column if exists status,
--   drop column if exists expected_amount, drop column if exists settled_at;
--   alter table public.events drop column if exists revenue_confirmed_at, drop column if exists revenue_confirmed_by;
