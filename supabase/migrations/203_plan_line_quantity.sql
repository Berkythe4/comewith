-- ============================================================
-- COME WITH — 203 planning: a line is QUANTITY x RATE, not a lump
--
-- WHY. The pricing model could say "$2,500 of ticket sales" but not "100 tickets
-- at $25". Those are the same number and a different amount of knowledge: the
-- first cannot be argued with, and the second can — you can look at 100 and say
-- we have never sold 100, or look at $25 and say the door is $30 now. Keith's
-- ask, exactly: quantity, then amount, then total, on every line, revenue and
-- cost alike. "1 production fee at $500 = $500" is worth writing as three
-- numbers even when the quantity is 1, because a reader can then see that it IS
-- one, rather than a lump that happens to round.
--
-- WHAT IT DOES NOT CHANGE. `quantity` defaults to 1, so every existing line
-- computes exactly what it computed yesterday. This migration cannot move a
-- forecast figure; it only gives the figures somewhere to be broken down.
--
-- HOW IT COMPOSES WITH `basis` (the one thing to get right here):
--
--   per_unit     total = quantity x amount
--                "2 DJs at $200" -> 2 x 200 = $400 per occurrence
--
--   per_scale    total = quantity x amount x scale
--                quantity is a MULTIPLIER ON TOP of the scale driver, so
--                "2 drinks a head at $4" is 2 x 4 x 43.5 attendance = $348.
--                The common case is quantity 1, which reads as "$25 a head".
--
--   pct_revenue  total = (amount / 100) x revenue.  Quantity is MEANINGLESS
--                against a percentage — "2 x 6% of revenue" is not a thing
--                anybody means — so it is pinned to 1 by constraint rather
--                than silently ignored by the views. A field the maths quietly
--                drops is worse than a field you cannot fill in wrong.
--
-- `unit_label` is the noun: "tickets", "DJs", "hours", "drinks". Nullable,
-- because plenty of lines have no natural noun, and a defaulted one ("units"
-- against a venue hire) is noise pretending to be information. Display only —
-- nothing computes from it.
--
-- Additive: one defaulted column, one nullable column, two views replaced with
-- the same column list. Per the destructive-vs-additive rule this may land
-- ahead of its UI — the deployed dashboard neither reads nor writes `quantity`
-- and keeps working unchanged until the new build ships.
-- ============================================================
begin;

alter table public.plan_offering_lines
  add column if not exists quantity   numeric(12,4) not null default 1,
  add column if not exists unit_label text;

alter table public.plan_offering_lines drop constraint if exists plan_offering_lines_quantity_check;
alter table public.plan_offering_lines add constraint plan_offering_lines_quantity_check
  check (quantity >= 0);

-- A percentage has no count. Pin it rather than ignore it.
alter table public.plan_offering_lines drop constraint if exists plan_offering_lines_pct_qty_check;
alter table public.plan_offering_lines add constraint plan_offering_lines_pct_qty_check
  check (basis <> 'pct_revenue' or quantity = 1);

comment on column public.plan_offering_lines.quantity is
  'How many, per one unit of the offering. per_unit: a plain count (2 DJs). '
  'per_scale: a multiplier ON TOP of the scale driver (2 drinks a head). '
  'pct_revenue: pinned to 1 by constraint — a percentage has no count.';
comment on column public.plan_offering_lines.unit_label is
  'The noun for the quantity — "tickets", "DJs", "hours". Display only; nothing '
  'computes from it. Null where a line has no natural noun.';

-- ---------------------------------------------------------------
-- Both halves of the maths, now x quantity
-- ---------------------------------------------------------------
create or replace view public.v_plan_offering_unit as
with flat as (
  select o.id as offering_id,
         coalesce(sum(case when l.direction = 'income' then
                        case l.basis when 'per_unit'  then l.quantity * l.amount
                                     when 'per_scale' then l.quantity * l.amount * o.default_scale
                                     else 0 end else 0 end), 0) as revenue_per_unit,
         coalesce(sum(case when l.direction = 'expense' and l.basis <> 'pct_revenue' then
                        case l.basis when 'per_unit'  then l.quantity * l.amount
                                     when 'per_scale' then l.quantity * l.amount * o.default_scale
                                     else 0 end else 0 end), 0) as cost_flat,
         coalesce(sum(case when l.direction = 'expense' and l.basis = 'pct_revenue'
                           then l.amount else 0 end), 0) as pct_rate,
         count(l.id)                                       as line_count,
         count(l.id) filter (where l.needs_review)         as unreviewed_lines,
         count(l.id) filter (where l.direction = 'income')  as income_lines,
         count(l.id) filter (where l.direction = 'expense') as expense_lines
    from public.plan_offerings o
    left join public.plan_offering_lines l
           on l.offering_id = o.id and l.deleted_at is null
   where o.deleted_at is null
   group by o.id
)
select o.id, o.key, o.label, o.ledger, o.creates_event, o.event_type, o.series,
       o.scale_label, o.default_scale, o.active, o.sort_order,
       round(f.revenue_per_unit, 2) as revenue_per_unit,
       round(f.cost_flat + (f.pct_rate / 100.0) * f.revenue_per_unit, 2) as cost_per_unit,
       round(f.revenue_per_unit - (f.cost_flat + (f.pct_rate / 100.0) * f.revenue_per_unit), 2)
         as contribution_per_unit,
       -- null, not 0%, when there is no revenue to take a percentage of
       case when f.revenue_per_unit > 0 then
         round(((f.revenue_per_unit - (f.cost_flat + (f.pct_rate / 100.0) * f.revenue_per_unit))
                / f.revenue_per_unit) * 100, 1) end as contribution_margin_pct,
       f.line_count, f.unreviewed_lines,
       (f.line_count = 0 or f.unreviewed_lines > 0) as provisional,
       -- Kept from 199: a missing side of the model must read as "no cost", not
       -- as a confident 100% margin (LEARNINGS §48).
       (f.income_lines  > 0) as has_revenue_model,
       (f.expense_lines > 0) as has_cost_model
  from public.plan_offerings o
  join flat f on f.offering_id = o.id
 where o.deleted_at is null;

comment on view public.v_plan_offering_unit is
  'What one unit of an offering earns, costs and contributes, from lines priced '
  'as quantity x rate. `provisional` is true while any line needs review; '
  'has_cost_model / has_revenue_model say whether a side of the model exists at '
  'all, so a missing cost side is never rendered as a 100% margin.';

create or replace view public.v_plan_monthly as
with vol as (
  select v.version_id, v.offering_id, v.period, v.units,
         coalesce(v.scale, o.default_scale) as scale,   -- null scale means "default", not zero
         o.ledger
    from public.plan_volumes v
    join public.plan_offerings o on o.id = v.offering_id and o.deleted_at is null
   where v.units > 0
),
-- Revenue per unit at THIS period's scale, so a pct_revenue cost follows a
-- December room that is twice the usual size.
rpu as (
  select vol.version_id, vol.offering_id, vol.period,
         coalesce(sum(case l.basis when 'per_unit'  then l.quantity * l.amount
                                   when 'per_scale' then l.quantity * l.amount * vol.scale
                                   else 0 end), 0) as revenue_per_unit
    from vol
    join public.plan_offering_lines l
      on l.offering_id = vol.offering_id and l.deleted_at is null
     and l.direction = 'income' and l.basis <> 'pct_revenue'
   group by 1, 2, 3
),
modelled as (
  select vol.version_id, vol.period, vol.ledger,
         case when l.direction = 'income' then 'revenue' else 'direct' end as section,
         l.category,
         sum(vol.units * case l.basis
               when 'per_unit'    then l.quantity * l.amount
               when 'per_scale'   then l.quantity * l.amount * vol.scale
               when 'pct_revenue' then (l.amount / 100.0) * coalesce(r.revenue_per_unit, 0)
             end) as amount
    from vol
    join public.plan_offering_lines l
      on l.offering_id = vol.offering_id and l.deleted_at is null
    left join rpu r on r.version_id  = vol.version_id
                   and r.offering_id = vol.offering_id
                   and r.period      = vol.period
   group by 1, 2, 3, 4, 5
),
-- Standing overhead. Only rows carrying a version: the 37 legacy rows stay out.
overhead as (
  select b.version_id, b.period, b.ledger,
         case when b.direction = 'income' then 'revenue' else 'indirect' end as section,
         b.category, sum(b.planned_amount) as amount
    from public.budget_lines b
   where b.version_id is not null and b.deleted_at is null
     and b.scope = 'period' and b.period is not null
   group by 1, 2, 3, 4, 5
),
rolled as (
  select version_id, period, ledger, section, category, sum(amount) as modelled_amount
    from (select * from modelled union all select * from overhead) z
   group by 1, 2, 3, 4, 5
),
keys as (
  select version_id, period, ledger, section, category from rolled
  union
  select version_id, period, ledger, section, category from public.plan_overrides
)
select k.version_id, k.period, k.ledger, k.section, k.category,
       round(coalesce(r.modelled_amount, 0), 2) as modelled_amount,
       o.amount                                 as override_amount,
       round(coalesce(o.amount, r.modelled_amount, 0), 2) as amount,
       (o.id is not null)                       as is_override
  from keys k
  left join rolled r
    on  r.version_id = k.version_id and r.period = k.period and r.ledger = k.ledger
    and r.section    = k.section    and r.category = k.category
  left join public.plan_overrides o
    on  o.version_id = k.version_id and o.period = k.period and o.ledger = k.ledger
    and o.section    = k.section    and o.category = k.category;

comment on view public.v_plan_monthly is
  'The forecast for a plan version: offering volumes x their lines (quantity x '
  'rate), plus standing overhead, with any typed-over cell REPLACING the modelled '
  'figure. is_override says which is which so the board can show what was '
  'asserted vs computed.';

-- Replacing a view re-grants nothing, but state it rather than inherit it:
-- these are financial-adjacent and stay anon-revoked (E1 / the 016-017 regression).
revoke select on public.v_plan_offering_unit from anon;
revoke select on public.v_plan_monthly       from anon;

commit;
