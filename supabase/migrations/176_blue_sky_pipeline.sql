-- ============================================================
-- COME WITH — 176 Blue Sky events, and upcoming shows that have no money on them
--
-- TWO GAPS, ONE SHAPE.
--
-- 1. There is no way to write down a gig you HOPE to book. Events are real or
--    they do not exist, so a year's plan lives in Keith's head and the forecast
--    only knows about period budget lines that are not attached to anything.
--
-- 2. v_event_money.missing_revenue only looks BACKWARDS — past events that
--    carry costs and no fee. An upcoming show with nothing on it is invisible
--    until it becomes a past show with nothing on it, which is too late to do
--    anything about.
--
-- BLUE SKY IS A STAGE, NOT A NEW TABLE. events.stage already permits 'idea',
-- and that is exactly what a speculative booking is. It needs two numbers to be
-- useful in a forecast:
--
--     expected_revenue — what it pays if it happens
--     confidence       — 0-100, how likely that is
--
-- weighted_revenue is the product. Ten $1,000 gigs at 30% is $3,000 of forecast,
-- not $10,000 — which is the whole point of writing them down rather than hoping.
--
-- A Blue Sky event is REPLACED by promoting it: change stage to 'confirmed' and
-- book real income against it. Or dropped: status 'cancelled'. Either way the
-- row survives, so the hit rate is measurable instead of anecdotal. Nothing here
-- touches the P&L — v_pl_monthly reads income and expenses, and a Blue Sky event
-- has neither until it becomes real.
-- ============================================================
begin;

alter table public.events
  add column if not exists expected_revenue numeric(12,2),
  add column if not exists confidence smallint;

alter table public.events
  drop constraint if exists events_confidence_range;
alter table public.events
  add constraint events_confidence_range
  check (confidence is null or (confidence >= 0 and confidence <= 100));

comment on column public.events.expected_revenue is
  'What this pays if it happens. Speculative — never counted as revenue.';
comment on column public.events.confidence is
  '0-100 likelihood. Multiplied into expected_revenue for the weighted pipeline.';

-- ---------------------------------------------------------------
-- The pipeline: everything not yet in the past, plus every Blue Sky idea
-- regardless of date, with what it is worth and whether anyone has said.
-- ---------------------------------------------------------------
create or replace view public.v_pipeline as
with booked as (
  select event_id,
         coalesce(sum(amount) filter (where status = 'received'), 0)  as received,
         coalesce(sum(coalesce(expected_amount, amount)) filter (where status <> 'received'), 0) as accrued
    from public.income
   where deleted_at is null
   group by event_id
), cost as (
  select event_id, coalesce(sum(amount), 0) as spent
    from public.expenses
   where deleted_at is null and event_id is not null
   group by event_id
)
select
  e.id                                   as event_id,
  e.name,
  e.series,
  e.type,
  e.event_date,
  e.status,
  coalesce(e.stage, 'planning')          as stage,
  (coalesce(e.stage, '') = 'idea')       as blue_sky,
  e.expected_revenue,
  e.confidence,
  round(coalesce(e.expected_revenue, 0) * coalesce(e.confidence, 0) / 100.0, 2) as weighted_revenue,
  coalesce(b.received, 0)                as booked_revenue,
  coalesce(b.accrued, 0)                 as accrued_revenue,
  coalesce(c.spent, 0)                   as spent_so_far,
  -- The working list: it has not happened yet, nobody has booked money against
  -- it, and nobody has said what it is expected to be worth.
  (e.event_date >= current_date
   and coalesce(b.received, 0) = 0
   and coalesce(b.accrued, 0) = 0
   and coalesce(e.expected_revenue, 0) = 0
   and e.status <> 'cancelled')          as needs_revenue_estimate
from public.events e
left join booked b on b.event_id = e.id
left join cost   c on c.event_id = e.id
where e.deleted_at is null
  and (e.event_date >= current_date or coalesce(e.stage, '') = 'idea');

revoke select on public.v_pipeline from anon;

commit;

-- DOWN:
--   drop view public.v_pipeline;
--   alter table public.events drop constraint events_confidence_range,
--     drop column expected_revenue, drop column confidence;
