-- ============================================================
-- COME WITH — 156 separate Dance Infusion, and track real cash
--
-- 1. DANCE INFUSION IS NOT COME WITH. It has its own bank account and every
--    penny of proceeds goes to the MS Society. Folding it into this P&L was
--    badly misleading: of the $13,049.83 the P&L called revenue, $12,497.94 was
--    DI (tickets 3,455.50 + sponsors 6,225 + donations 2,817.44). Come With's own
--    revenue is about $552. `ledger` splits them so neither flatters the other.
--
--    Derived from the event's series, because that is already how DI is
--    identified here, with an explicit column so a non-event DI cost can be
--    marked by hand.
--
-- 2. CASH IS NOT COST. The $5,000 float only moves when money leaves the PayPal
--    account or the business bank. Everything else — spend on the personal card,
--    the hand-kept ledger — is a real cost but never touched that $5,000. The old
--    v_cash_position subtracted ALL business-flagged spend and produced
--    -$10,839.81, which was never a bank balance.
--
--    `payment_method` cannot answer this: it holds only NULL (185) and
--    'Cash/Card' (62). So cash_source is explicit, and Jennifer sets it on push
--    from what it already knows (paypal_ingest / simplifi_cw_mirror / ledger).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Which set of books a row belongs to
-- ---------------------------------------------------------------
alter table public.expenses add column if not exists ledger text
  not null default 'come_with';
alter table public.expenses drop constraint if exists expenses_ledger_check;
alter table public.expenses add constraint expenses_ledger_check
  check (ledger in ('come_with', 'dance_infusion'));

alter table public.income add column if not exists ledger text
  not null default 'come_with';
alter table public.income drop constraint if exists income_ledger_check;
alter table public.income add constraint income_ledger_check
  check (ledger in ('come_with', 'dance_infusion'));

comment on column public.expenses.ledger is
  'Which business this belongs to. Dance Infusion has its own bank account and '
  'donates all proceeds, so it must never be mixed into the Come With P&L.';

-- Backfill from the event series — the existing source of truth for DI.
update public.expenses e set ledger = 'dance_infusion'
  from public.events ev
 where ev.id = e.event_id and ev.series ilike '%dance infusion%' and e.ledger <> 'dance_infusion';

update public.income i set ledger = 'dance_infusion'
  from public.events ev
 where ev.id = i.event_id and ev.series ilike '%dance infusion%' and i.ledger <> 'dance_infusion';

create index if not exists idx_expenses_ledger on public.expenses(ledger);

-- ---------------------------------------------------------------
-- 2. Where the money physically moved
-- ---------------------------------------------------------------
alter table public.expenses add column if not exists cash_source text;
alter table public.expenses drop constraint if exists expenses_cash_source_check;
alter table public.expenses add constraint expenses_cash_source_check
  check (cash_source is null or cash_source in ('paypal', 'bank', 'personal', 'other'));

comment on column public.expenses.cash_source is
  'Which pot the money physically left. Only paypal and bank draw down the '
  'business cash float; personal is Keith''s card (capital, not business cash). '
  'NULL means nobody has said yet.';

create index if not exists idx_expenses_cash_source on public.expenses(cash_source);

-- Anything already flagged as personally funded is, by definition, personal cash.
update public.expenses set cash_source = 'personal'
 where deleted_at is null and funded_by = 'owner' and cash_source is null;

-- ---------------------------------------------------------------
-- 3. P&L, per ledger
-- ---------------------------------------------------------------
create or replace view public.v_pl_monthly as
with rev as (
  select to_char(coalesce(t.purchased_at::date, e.event_date, t.created_at::date), 'YYYY-MM') as period,
         'Ticket sales'::text as category, sum(t.amount_paid) as revenue, 0::numeric as cost, count(*) as n,
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end as ledger
    from public.ticketing t left join public.events e on e.id = t.event_id
   where t.amount_paid is not null group by 1, 6
  union all
  select to_char(coalesce(e.event_date, s.created_at::date), 'YYYY-MM'), 'Sponsorship',
         sum(s.cash_amount), 0, count(*),
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
    from public.sponsorships s left join public.events e on e.id = s.event_id
   where s.status <> 'cancelled' and s.cash_amount is not null group by 1, 6
  union all
  select to_char(coalesce(d.date, e.event_date, d.created_at::date), 'YYYY-MM'), 'Donations',
         sum(d.amount), 0, count(*),
         case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
    from public.third_party_donations d left join public.events e on e.id = d.event_id
   where d.amount is not null group by 1, 6
  union all
  select to_char(i.date, 'YYYY-MM'), coalesce(nullif(i.category, ''), 'Other income'),
         sum(i.amount), 0, count(*), i.ledger
    from public.income i where i.deleted_at is null group by 1, 2, 6
),
cost as (
  select to_char(date, 'YYYY-MM') as period, coalesce(nullif(category,''), 'Uncategorised') as category,
         0::numeric as revenue, sum(amount) as cost, count(*) as n,
         (event_id is not null) as is_direct, ledger
    from public.expenses where deleted_at is null
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

create or replace view public.v_pl_summary as
-- `ledger` must be APPENDED, not inserted: create-or-replace can add trailing
-- columns but never reorder existing ones.
select period,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0), 2)  as revenue,
       round(coalesce(sum(cost)    filter (where section = 'direct'), 0), 2)   as direct_cost,
       round(coalesce(sum(cost)    filter (where section = 'indirect'), 0), 2) as indirect_cost,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0)
           - coalesce(sum(cost)    filter (where section = 'direct'), 0), 2)   as gross_profit,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0)
           - coalesce(sum(cost)    filter (where section in ('direct','indirect')), 0), 2) as net_profit,
       ledger
  from public.v_pl_monthly
 group by period, ledger;

revoke select on public.v_pl_summary from anon;

-- ---------------------------------------------------------------
-- 4. Cash: the float, and only what actually moves it
-- ---------------------------------------------------------------
-- Columns are RENAMED here (spent_by_business -> spent_from_float), which
-- create-or-replace cannot do. Nothing reads this view but the dashboard, so
-- dropping it is safe.
drop view if exists public.v_cash_position;
create view public.v_cash_position as
with cap as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions
   where deleted_at is null and kind = 'cash'
), inflow as (
  -- Come With revenue only. DI money lands in DI's own account.
  select coalesce(sum(revenue), 0) as amt from public.v_pl_monthly
   where section = 'revenue' and ledger = 'come_with'
), outflow as (
  select coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source in ('paypal', 'bank')
), unknown_src as (
  select count(*) as n, coalesce(sum(amount), 0) as amt from public.expenses
   where deleted_at is null and ledger = 'come_with' and cash_source is null
)
select cap.amt                              as capital_in,
       inflow.amt                           as revenue_in,
       outflow.amt                          as spent_from_float,
       cap.amt + inflow.amt - outflow.amt   as cash_reserve,
       unknown_src.n                        as unknown_source_rows,
       unknown_src.amt                      as unknown_source_amount
  from cap, inflow, outflow, unknown_src;

revoke select on public.v_cash_position from anon;

commit;

-- DOWN: restore 154's views, then
--   alter table public.expenses drop column if exists ledger, drop column if exists cash_source;
--   alter table public.income drop column if exists ledger;
