-- ============================================================
-- COME WITH — 021 IG FOLLOWINGS  (additive + migrate generic key)
-- Track three separate IG followings instead of one generic metric.
-- Deactivate the generic instagram.followers (and its same-day dup) so it
-- drops out of v_kpi_targets_current — history is kept, not deleted.
-- Seed sensible starter targets; Keith adjusts from the UI.
-- ============================================================
begin;

-- retire the generic metric (both active rows)
update public.kpi_targets set active = false where metric_key = 'instagram.followers';

-- three tracked followings
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit) values
  ('instagram.followers.comewith',      'audience',       'IG — Come With',      5000, 'gte', ''),
  ('instagram.followers.berky',         'audience',       'IG — Berky',          3000, 'gte', ''),
  ('instagram.followers.danceinfusion', 'dance_infusion', 'IG — Dance Infusion', 2000, 'gte', '');

commit;
