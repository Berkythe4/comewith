-- =============================================================================
-- 050_expense_event_na.sql  (additive)
-- expenses.event_na — "not applicable to events" (business overhead / capital that
-- will never belong to a single event: software, equipment, etc). Lets the Expenses
-- tab show a clean "truly unassigned = needs an event" list separate from overhead.
-- Defaults Software/Equipment (unassigned) to N/A. Admin-only via existing RLS.
-- =============================================================================
begin;

alter table public.expenses add column if not exists event_na boolean not null default false;

update public.expenses set event_na = true
  where deleted_at is null and event_id is null and event_na = false
    and category in ('Software', 'Equipment');

commit;

-- DOWN: alter table public.expenses drop column event_na;
