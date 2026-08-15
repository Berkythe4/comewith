-- =============================================================================
-- 143_kpi_last_event_cards.sql   (Strategy rebuild -- Phase 2: card the
-- last-event metrics)
--
-- 142 computed these but deliberately did not card them: a value in
-- v_kpi_computed is inert until a kpi_targets row exists, which kept the
-- deployed dashboard pixel-identical while the data layer changed underneath.
-- The new renderer leads each category with them, so they need rows now.
--
-- INTERIM EFFECT, on purpose: until the Phase 2 dashboard.html is merged, the
-- DEPLOYED board picks these up as six additional cards inside the existing
-- Parties / Dance Infusion / Content groups. Additive and readable -- the old
-- renderer resolves all three workstreams and colours them correctly -- just
-- redundant next to the lifetime averages until the merge lands.
--
-- Targets mirror the existing per-event rows so nothing is invented here:
-- parties.net_pl 0, parties.sell_through 100, di.cost_to_raise 0.50,
-- di.raised_per_event 6000, di.attendance 250. They are all editable from the
-- dashboard ("Edit target") -- SQL is not where these get tuned.
--
-- Labels are PLAIN ASCII: an em-dash mojibakes through the Management API.
-- =============================================================================
begin;

insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('parties.net_pl_last',       'parties',        'Net P&L - last party',        0,    'gte', '$', current_date, true),
  ('parties.sell_through_last', 'parties',        'Sell-through - last party',   100,  'gte', '%', current_date, true),
  ('di.cost_to_raise_last',     'dance_infusion', 'Cost to raise $1 - last',     0.50, 'lte', '$', current_date, true),
  ('di.raised_last',            'dance_infusion', 'Raised - last event',         6000, 'gte', '$', current_date, true),
  ('di.attendance_last',        'dance_infusion', 'Attendance - last event',     250,  'gte', '',  current_date, true),
  -- youtube.avg_views is lifetime views / all videos and cannot move; this is
  -- the last 5 uploads. Target 500 rather than that card's 4000: at ~103 avg
  -- views on recent uploads, 4000 is not a target, it is noise.
  ('content.avg_views_recent',  'content',        'Avg views - last 5 uploads',  500,  'gte', '',  current_date, true)
on conflict do nothing;

commit;

-- Post-apply: all six resolve through v_kpi_dashboard, and no metric_key ends
-- up with two active target rows (142 cleaned the existing duplicates).
--   select metric_key, current_value, prior_value, prior_basis, source_kind
--     from public.v_kpi_dashboard
--    where metric_key in ('parties.net_pl_last','parties.sell_through_last',
--                         'di.cost_to_raise_last','di.raised_last',
--                         'di.attendance_last','content.avg_views_recent');
--
-- DOWN: update public.kpi_targets set active = false where metric_key in (...);
