-- =============================================================================
-- 075_audit_survey_responses.sql
-- Put survey_responses under the standard audit_trigger_function() so every
-- submission — and, crucially, any DELETE — lands in audit_log with the acting
-- user + timestamp. Motivation: when asked to prove "did anyone answer before?"
-- we could confirm the table currently holds only the test response, but survey
-- tables had NO audit trigger, so an insert-then-delete would have left no trace.
-- This closes that gap going forward. Reuses the function from 010; additive only.
-- (survey_answers left out as high-volume/low-signal per the 059 policy —
-- survey_responses is the authoritative "a response exists" record.)
-- =============================================================================
begin;
drop trigger if exists audit_survey_responses on public.survey_responses;
create trigger audit_survey_responses
  after insert or update or delete on public.survey_responses
  for each row execute function public.audit_trigger_function();
commit;
-- POST: inserts/updates/deletes on survey_responses now appear in audit_log
-- (master-only read), visible in Users → Activity. No change to RLS or grants.
