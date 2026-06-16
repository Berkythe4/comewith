-- =============================================================================
-- 036_contract_outreach_template.sql  —  Sprint 4: contract-related outreach task
-- Adds a venue-contract touch-base to the standard checklist so the generated chain
-- includes a contract task that auto-assigns to the venue's booking contact (and is
-- editable in the Templates tab like any other). Additive seed; idempotent.
-- =============================================================================
insert into public.task_templates (event_type, title, default_offset_days, phase, gear_applicability, target_function, sort_order)
select et.t, 'Finalize & sign venue contract', -21, 'planning', 'both', 'venue:booking', 3
from (values ('party'), ('dance_infusion')) as et(t)
on conflict (event_type, title) do nothing;
