-- ============================================================
-- COME WITH — 198 planning views + seed from the legacy budget
--
-- 197 built the tables. This turns them into the four things the board reads:
--
--   v_plan_offering_unit  unit economics — what one party / gig / rental earns,
--                         costs and contributes. The lever panel is this view.
--   v_plan_monthly        the forecast: volumes x offering lines, plus standing
--                         overhead, with typed-over cells replacing the model.
--   v_plan_vs_actual      the same, joined to v_pl_monthly on (period, category)
--                         — the join the legacy budget rows could never make.
--   v_event_contribution  BACKWARD: what each event that actually happened
--                         contributed, against what the model said it would.
--                         This is how the model gets better instead of staying
--                         a guess forever.
--
-- SIGN CONVENTION. Costs are POSITIVE everywhere here, matching expenses.amount
-- and v_pl_monthly (149 chose that and flipping it would rewrite every view).
-- So `variance` is always actual minus plan, and whether that is GOOD depends on
-- the section — which is why `favourable` is computed here rather than left to
-- each caller to get wrong: over on revenue is good, over on cost is not.
--
-- NULL IS NOT ZERO (LEARNINGS §23). Margin percentages return null when there is
-- no revenue to divide by, and variance_pct returns null against a zero plan.
-- "0%" and "cannot be computed" are opposite claims and the board must not
-- render them the same.
--
-- THE SEED. Amounts come from rows that already exist — the 37 legacy
-- budget_lines figures, and real average ticket price and attendance computed
-- from `ticketing`. Nothing is invented. What could NOT be derived is which P&L
-- category a legacy lump belongs to (the $1,200 against "Come With Party #1"
-- covers venue, talent and marketing as one number), so every line seeded from a
-- lump carries needs_review = true and the offering reads as PROVISIONAL until a
-- human splits it. LEARNINGS §26: a placeholder that feeds a sum stops being a
-- placeholder and becomes invented evidence.
-- ============================================================
begin;

-- pct_revenue amounts are PERCENTS (6 = 6%), not fractions. Storing 0.06 would
-- read as six cents to anyone looking at the table.
alter table public.plan_offering_lines drop constraint if exists plan_offering_lines_pct_range_check;
alter table public.plan_offering_lines add constraint plan_offering_lines_pct_range_check
  check (basis <> 'pct_revenue' or (amount >= 0 and amount <= 100));

comment on column public.plan_offering_lines.amount is
  'per_unit: dollars per occurrence. per_scale: dollars per unit of scale (per '
  'head, per item). pct_revenue: a PERCENT (6 = 6%), never a fraction.';

-- ---------------------------------------------------------------
-- 1. Unit economics — the lever panel
-- ---------------------------------------------------------------
create or replace view public.v_plan_offering_unit as
with flat as (
  select o.id as offering_id,
         coalesce(sum(case when l.direction = 'income' then
                        case l.basis when 'per_unit'  then l.amount
                                     when 'per_scale' then l.amount * o.default_scale
                                     else 0 end else 0 end), 0) as revenue_per_unit,
         coalesce(sum(case when l.direction = 'expense' and l.basis <> 'pct_revenue' then
                        case l.basis when 'per_unit'  then l.amount
                                     when 'per_scale' then l.amount * o.default_scale
                                     else 0 end else 0 end), 0) as cost_flat,
         coalesce(sum(case when l.direction = 'expense' and l.basis = 'pct_revenue'
                           then l.amount else 0 end), 0) as pct_rate,
         count(l.id)                                  as line_count,
         count(l.id) filter (where l.needs_review)    as unreviewed_lines
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
       (f.line_count = 0 or f.unreviewed_lines > 0) as provisional
  from public.plan_offerings o
  join flat f on f.offering_id = o.id
 where o.deleted_at is null;

comment on view public.v_plan_offering_unit is
  'What one unit of an offering earns, costs and contributes. `provisional` is '
  'true while any line still needs review — the board must not present a model '
  'built on a guessed category as settled.';

-- ---------------------------------------------------------------
-- 2. The forecast
-- ---------------------------------------------------------------
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
         coalesce(sum(case l.basis when 'per_unit'  then l.amount
                                   when 'per_scale' then l.amount * vol.scale
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
               when 'per_unit'    then l.amount
               when 'per_scale'   then l.amount * vol.scale
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
  'The forecast for a plan version: offering volumes x their lines, plus standing '
  'overhead, with any typed-over cell REPLACING the modelled figure. is_override '
  'says which is which so the board can show what was asserted vs computed.';

-- ---------------------------------------------------------------
-- 3. Plan vs actual
-- ---------------------------------------------------------------
create or replace view public.v_plan_vs_actual as
with actual as (
  select period, ledger, section, category,
         sum(case when section = 'revenue' then revenue else cost end) as actual_amount
    from public.v_pl_monthly
   group by 1, 2, 3, 4
),
plan as (
  select version_id, period, ledger, section, category, amount as plan_amount
    from public.v_plan_monthly
),
-- Actuals only enter for periods the plan actually covers; otherwise every
-- version would inherit all 23 months of history as "unplanned".
span as (select distinct version_id, period from plan),
keys as (
  select version_id, period, ledger, section, category from plan
  union
  select s.version_id, a.period, a.ledger, a.section, a.category
    from actual a join span s on s.period = a.period
)
select k.version_id, k.period, k.ledger, k.section, k.category,
       round(coalesce(p.plan_amount, 0), 2)   as plan_amount,
       round(coalesce(a.actual_amount, 0), 2) as actual_amount,
       round(coalesce(a.actual_amount, 0) - coalesce(p.plan_amount, 0), 2) as variance,
       -- null against a zero plan: a percentage of nothing is not 0%
       case when coalesce(p.plan_amount, 0) <> 0 then
         round(((coalesce(a.actual_amount, 0) - p.plan_amount) / abs(p.plan_amount)) * 100, 1)
       end as variance_pct,
       -- over on revenue is good; over on cost is not. Computed once, here.
       case when k.section = 'revenue'
            then coalesce(a.actual_amount, 0) >= coalesce(p.plan_amount, 0)
            else coalesce(a.actual_amount, 0) <= coalesce(p.plan_amount, 0)
       end as favourable,
       (p.plan_amount is null)   as unplanned,
       (a.actual_amount is null) as no_actual_yet
  from keys k
  left join plan p
    on  p.version_id = k.version_id and p.period = k.period and p.ledger = k.ledger
    and p.section    = k.section    and p.category = k.category
  left join actual a
    on  a.period  = k.period  and a.ledger   = k.ledger
    and a.section = k.section and a.category = k.category;

comment on view public.v_plan_vs_actual is
  'Forecast against actuals on (period, ledger, section, category). This is the '
  'join the legacy budget_lines rows could never make, because they stored the '
  'unit NAME in category and no P&L category is called "DJ Gig #1".';

-- ---------------------------------------------------------------
-- 4. Backward: what each event actually contributed
-- ---------------------------------------------------------------
-- The forward model is only as good as its last comparison to reality. This is
-- the same contribution arithmetic as v_plan_offering_unit, run on what
-- happened, so `vs_model` says how wrong the model was per event.
create or replace view public.v_event_contribution as
select m.event_id, m.name, m.series, m.event_date, m.ledger, m.status,
       to_char(m.event_date, 'YYYY-MM') as period,
       e.type as event_type,
       o.key   as offering_key,
       o.label as offering_label,
       round(coalesce(m.revenue, 0), 2)  as revenue,
       round(coalesce(m.expenses, 0), 2) as direct_cost,
       round(coalesce(m.revenue, 0) - coalesce(m.expenses, 0), 2) as contribution,
       case when coalesce(m.revenue, 0) > 0 then
         round(((coalesce(m.revenue, 0) - coalesce(m.expenses, 0)) / m.revenue) * 100, 1)
       end as contribution_margin_pct,
       e.total_attendance,
       case when coalesce(e.total_attendance, 0) > 0 then
         round((coalesce(m.revenue, 0) - coalesce(m.expenses, 0)) / e.total_attendance, 2)
       end as contribution_per_head,
       u.contribution_per_unit as modelled_contribution,
       case when u.contribution_per_unit is not null then
         round((coalesce(m.revenue, 0) - coalesce(m.expenses, 0)) - u.contribution_per_unit, 2)
       end as vs_model,
       m.upcoming
  from public.v_event_money m
  join public.events e on e.id = m.event_id and e.deleted_at is null
  -- Deterministic pick by sort_order rather than a unique constraint on
  -- event_type: two party offerings (small room / big room) must stay legal.
  left join lateral (
    select po.id, po.key, po.label
      from public.plan_offerings po
     where po.deleted_at is null and po.active and po.creates_event
       and po.event_type = e.type and po.ledger = m.ledger
     order by po.sort_order, po.key
     limit 1
  ) o on true
  left join public.v_plan_offering_unit u on u.id = o.id;

comment on view public.v_event_contribution is
  'What an event that actually happened earned, cost and contributed, next to '
  'what the offering model predicted. vs_model is how the pricing model gets '
  'corrected instead of staying a guess.';

-- ---------------------------------------------------------------
-- 5. Anon stays out. E1 discipline / the 016-017 regression.
-- ---------------------------------------------------------------
revoke select on public.v_plan_offering_unit from anon;
revoke select on public.v_plan_monthly       from anon;
revoke select on public.v_plan_vs_actual     from anon;
revoke select on public.v_event_contribution from anon;
revoke select on public.plan_versions        from anon;
revoke select on public.plan_offerings       from anon;
revoke select on public.plan_offering_lines  from anon;
revoke select on public.plan_volumes         from anon;
revoke select on public.plan_overrides       from anon;

commit;

-- DOWN:
--   drop view if exists public.v_event_contribution, public.v_plan_vs_actual,
--                       public.v_plan_monthly, public.v_plan_offering_unit;
