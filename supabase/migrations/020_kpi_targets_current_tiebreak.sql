-- ============================================================
-- COME WITH — 020 v_kpi_targets_current: deterministic tiebreak  (corrective)
--
-- Bug: v_kpi_targets_current used DISTINCT ON (metric_key) ORDER BY metric_key,
-- effective_date DESC with NO tiebreaker. Two active targets for the same metric
-- on the SAME effective_date (e.g. an Edit-Target update made on the seed date)
-- resolved arbitrarily, so the edit saved a row but the card kept the old value
-- ("Update existing target does not work").
--
-- Fix: add updated_at DESC, id DESC so the most recently saved row wins.
-- Output columns unchanged -> CREATE OR REPLACE is safe; v_kpi_dashboard depends
-- on this view and is unaffected. Not a financial view (no anon revoke needed).
-- ============================================================
begin;

create or replace view public.v_kpi_targets_current as
select distinct on (metric_key)
  metric_key, workstream, label, target_value, comparison, unit, effective_date
from public.kpi_targets
where active
order by metric_key, effective_date desc, updated_at desc, id desc;

commit;
