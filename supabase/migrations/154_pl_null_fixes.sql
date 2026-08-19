-- ============================================================
-- COME WITH — 154 P&L: undated tickets, and profit lines that read 0 for NULL
--
-- Two defects visible the moment 153 put real revenue on screen.
--
-- 1. THREE TICKET SALES ($550) HAVE NO purchased_at, so they landed in a NULL
--    month — real revenue, banked, sitting outside every period. Fall back to the
--    event's date, then to created_at. A ticket always belongs to an event and an
--    event always has a date, so this cannot silently drop money again.
--
-- 2. `sum(...) filter (...)` returns NULL, not 0, when nothing matches. So any
--    month with cost but no revenue produced NULL gross and NULL net, which the
--    UI rendered as 0.00 — the most misleading possible value, since those are
--    exactly the months that lost the most. 2026-08 showed net 0.00 when it was
--    -1,665.55. Every aggregate is now coalesced.
-- ============================================================
begin;

create or replace view public.v_pl_monthly as
with rev as (
  -- coalesce chain: when a ticket was bought, else when the event happened,
  -- else when the row was created. Never NULL.
  select to_char(coalesce(t.purchased_at::date, e.event_date, t.created_at::date), 'YYYY-MM') as period,
         'Ticket sales'::text as category,
         sum(t.amount_paid) as revenue, 0::numeric as cost, count(*) as n
    from public.ticketing t
    left join public.events e on e.id = t.event_id
   where t.amount_paid is not null group by 1
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
cost as (
  select to_char(date, 'YYYY-MM') as period, coalesce(nullif(category,''), 'Uncategorised') as category,
         0::numeric as revenue, sum(amount) as cost, count(*) as n,
         (event_id is not null) as is_direct
    from public.expenses where deleted_at is null
   group by 1, 2, 6
)
select period, category,
       'revenue'::text as bucket,
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

-- Every filtered aggregate coalesced. A month with no revenue now reports its
-- loss instead of reporting nothing.
create or replace view public.v_pl_summary as
select period,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0), 2)  as revenue,
       round(coalesce(sum(cost)    filter (where section = 'direct'), 0), 2)   as direct_cost,
       round(coalesce(sum(cost)    filter (where section = 'indirect'), 0), 2) as indirect_cost,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0)
           - coalesce(sum(cost)    filter (where section = 'direct'), 0), 2)   as gross_profit,
       round(coalesce(sum(revenue) filter (where section = 'revenue'), 0)
           - coalesce(sum(cost)    filter (where section in ('direct','indirect')), 0), 2) as net_profit
  from public.v_pl_monthly
 group by period;

revoke select on public.v_pl_summary from anon;

commit;

-- DOWN: revert to 153's definitions.
