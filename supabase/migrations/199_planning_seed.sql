-- ============================================================
-- COME WITH — 199 seed the planner from the legacy budget
--
-- Creates the working plan round and six offerings, derived from rows that
-- already exist. Every figure below is COMPUTED from the 37 legacy
-- budget_lines rows, from `ticketing`, or from `events.total_attendance` — the
-- migration hardcodes the STRUCTURE (which unit is which event type) and
-- derives the NUMBERS, so it stays correct if the legacy rows are corrected
-- before it runs.
--
-- ---------------------------------------------------------------------------
-- WHAT COULD NOT BE DERIVED, AND WHAT WAS DONE ABOUT IT
-- ---------------------------------------------------------------------------
-- 1. CATEGORY. The legacy rows put the unit name in `category` ("DJ Gig #1"),
--    so the P&L category each lump belongs to is unknown. The mapping below is
--    a best guess and every seeded line carries needs_review = true, which
--    makes the offering read as PROVISIONAL on the board.
--
-- 2. PER-HEAD TICKET PRICING FOR PARTIES. There is no paid ticket row against
--    any event of type 'party' — priced ticketing exists only for Dance
--    Infusion. So an attendance-driven ticket line CANNOT be derived for
--    parties without inventing a price, and inventing one would feed straight
--    into every forecast number (LEARNINGS §26). Party revenue is therefore
--    seeded as the flat per-occurrence figure Keith actually budgeted, and
--    `default_scale` is set from real average attendance so the lever is ready
--    the moment a ticket price is entered — but it moves no money until then.
--
-- 3. ZERO IS NOT SEEDED AS A LINE. Equipment Rental and Event Production have
--    no expense in the legacy budget, and Artist Showcase has no income. A
--    $0.00 line asserts "this costs nothing"; a missing line says "not modelled
--    yet". They are opposite claims and only the second one is true, so no line
--    is written. v_plan_offering_unit reports has_cost_model / has_revenue_model
--    (added below) so the board can say which side is missing instead of
--    printing a confident 100% margin.
--
-- The legacy rows are NOT touched. They keep version_id null, stay readable as
-- history, and are invisible to every planner view.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 0. The unit view learns to say which side of the model is missing
-- ---------------------------------------------------------------
-- create-or-replace can append trailing columns but never reorder existing
-- ones, so these two go on the end.
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
       case when f.revenue_per_unit > 0 then
         round(((f.revenue_per_unit - (f.cost_flat + (f.pct_rate / 100.0) * f.revenue_per_unit))
                / f.revenue_per_unit) * 100, 1) end as contribution_margin_pct,
       f.line_count, f.unreviewed_lines,
       (f.line_count = 0 or f.unreviewed_lines > 0) as provisional,
       (f.income_lines  > 0) as has_revenue_model,
       (f.expense_lines > 0) as has_cost_model
  from public.plan_offerings o
  join flat f on f.offering_id = o.id
 where o.deleted_at is null;

revoke select on public.v_plan_offering_unit from anon;

comment on view public.v_plan_offering_unit is
  'What one unit of an offering earns, costs and contributes. `provisional` is '
  'true while any line needs review; has_cost_model / has_revenue_model say '
  'whether a side of the model exists at all, so a missing cost side is never '
  'rendered as a 100% margin.';

-- ---------------------------------------------------------------
-- 1. The working round
-- ---------------------------------------------------------------
insert into public.plan_versions (label, status, horizon_months, basis_period, notes)
select 'Working plan', 'working', 6, to_char(current_date, 'YYYY-MM'),
       'Seeded by migration 199 from the 37 legacy budget_lines rows.'
 where not exists (select 1 from public.plan_versions where status = 'working');

-- ---------------------------------------------------------------
-- 2. What each legacy unit was worth, derived
-- ---------------------------------------------------------------
create temporary table _legacy_unit on commit drop as
select regexp_replace(b.category, '\s*#\s*\d+.*$', '') as unit_name,
       b.category                                      as orig,
       b.period,
       sum(case when b.direction = 'income'  then b.planned_amount else 0 end) as inc,
       sum(case when b.direction = 'expense' then b.planned_amount else 0 end) as exp
  from public.budget_lines b
 where b.version_id is null and b.deleted_at is null
   and b.scope = 'period' and b.period is not null
   and b.category ~ '#\s*\d+'
 group by 1, 2, 3;

-- STRUCTURE is declared (which unit is which event type, and the best guess at
-- its P&L category); NUMBERS are derived from _legacy_unit below.
create temporary table _unit_map on commit drop as
select * from (values
  ('Come With Party',   'party',      'Come With Party',      'come_with',      true,  'party',          'Come With Parties',  'Paid attendance', 'Ticket sales',     'Venue',       10),
  ('DJ Gig',            'dj_booking', 'DJ Booking',           'come_with',      true,  'gig',            'Bookings',           'Bookings',        'Production fee',   'Contractors', 20),
  ('Event Production',  'production', 'Event Production',     'come_with',      true,  'production',     'Come With Production','Productions',    'Production fee',   'Contractors', 30),
  ('Equipment Rental',  'rental',     'Equipment Rental',     'come_with',      false, null,             null,                 'Rentals',         'Equipment rental', 'Equipment',   40),
  ('Artist Showcase',   'showcase',   'Artist Showcase',      'come_with',      true,  'showcase',       'Content Creation',   'Attendance',      'Other income',     'Operations',  50),
  ('Dance Infusion',    'di_event',   'Dance Infusion Event', 'dance_infusion', true,  'dance_infusion', 'Dance Infusion',     'Paid attendance', 'Ticket sales',     'Production',  60)
) as t(unit_name, key, label, ledger, creates_event, event_type, series,
       scale_label, income_category, expense_category, sort_order);

-- ---------------------------------------------------------------
-- 3. Offerings
-- ---------------------------------------------------------------
-- default_scale: real PAID attendance where priced tickets exist, otherwise
-- real recorded attendance, otherwise 1. Never a made-up number.
insert into public.plan_offerings
  (key, label, ledger, creates_event, event_type, series, scale_label, default_scale, sort_order, notes)
select m.key, m.label, m.ledger, m.creates_event, m.event_type, m.series, m.scale_label,
       coalesce(
         (select round(avg(h.heads), 2) from (
            select sum(coalesce(tk.quantity, 1)) as heads
              from public.ticketing tk
              join public.events e2 on e2.id = tk.event_id and e2.deleted_at is null
             where e2.type = m.event_type and tk.amount_paid > 0
             group by tk.event_id) h),
         (select round(avg(e3.total_attendance), 2) from public.events e3
           where e3.deleted_at is null and e3.type = m.event_type
             and e3.total_attendance is not null),
         1),
       m.sort_order,
       'Seeded by 199 from the legacy budget line "' || m.unit_name || '". '
       || 'Amounts are Keith''s own figures; the P&L category is a guess and is '
       || 'flagged for review.'
  from _unit_map m
 where exists (select 1 from _legacy_unit l where l.unit_name = m.unit_name)
on conflict (key) do nothing;

-- ---------------------------------------------------------------
-- 4. Lines — one income, one expense, only where a figure exists
-- ---------------------------------------------------------------
insert into public.plan_offering_lines
  (offering_id, direction, category, label, basis, amount, needs_review, sort_order)
select o.id, 'income', m.income_category, m.unit_name || ' revenue', 'per_unit',
       round(avg(l.inc), 2), true, 10
  from _legacy_unit l
  join _unit_map m       on m.unit_name = l.unit_name
  join public.plan_offerings o on o.key = m.key
 group by o.id, m.income_category, m.unit_name
having round(avg(l.inc), 2) > 0
   and not exists (select 1 from public.plan_offering_lines x
                    where x.offering_id = o.id and x.direction = 'income');

insert into public.plan_offering_lines
  (offering_id, direction, category, label, basis, amount, needs_review, sort_order)
select o.id, 'expense', m.expense_category, m.unit_name || ' cost', 'per_unit',
       round(avg(l.exp), 2), true, 20
  from _legacy_unit l
  join _unit_map m       on m.unit_name = l.unit_name
  join public.plan_offerings o on o.key = m.key
 group by o.id, m.expense_category, m.unit_name
having round(avg(l.exp), 2) > 0
   and not exists (select 1 from public.plan_offering_lines x
                    where x.offering_id = o.id and x.direction = 'expense');

-- ---------------------------------------------------------------
-- 5. Volumes — how many of each, per month, exactly as budgeted
-- ---------------------------------------------------------------
insert into public.plan_volumes (version_id, offering_id, period, units)
select v.id, o.id, l.period, count(distinct l.orig)
  from _legacy_unit l
  join _unit_map m             on m.unit_name = l.unit_name
  join public.plan_offerings o on o.key = m.key
  cross join (select id from public.plan_versions where status = 'working' limit 1) v
 group by v.id, o.id, l.period
on conflict (version_id, offering_id, period) do nothing;

-- ---------------------------------------------------------------
-- 6. Standing overhead into the working round
-- ---------------------------------------------------------------
-- Marketing and Software already carry REAL P&L categories, so unlike the unit
-- rows they join to actuals correctly and are copied across as they stand.
insert into public.budget_lines
  (scope, period, category, direction, planned_amount, ledger, version_id, label, notes)
select 'period', b.period, b.category, b.direction, b.planned_amount, 'come_with', v.id,
       b.category || ' (standing)',
       'Copied into the working plan by 199 from the legacy budget row.'
  from public.budget_lines b
  cross join (select id from public.plan_versions where status = 'working' limit 1) v
 where b.version_id is null and b.deleted_at is null
   and b.scope = 'period' and b.period is not null
   and b.category !~ '#\s*\d+'
   and not exists (
     select 1 from public.budget_lines x
      where x.version_id = v.id and x.period = b.period
        and x.category = b.category and x.direction = b.direction);

commit;

-- DOWN:
--   delete from public.budget_lines where version_id is not null;
--   delete from public.plan_volumes;
--   delete from public.plan_offering_lines;
--   delete from public.plan_offerings;
--   delete from public.plan_versions;
