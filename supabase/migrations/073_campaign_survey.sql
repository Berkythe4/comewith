-- =============================================================================
-- 073_campaign_survey.sql
-- Attach a survey to an email campaign. When the campaign sends, each recipient
-- gets a PERSONAL tokenized survey link in the email (responses tag to them +
-- the survey's event). Lets the impact-report email + its feedback survey go out
-- together as one send.
-- =============================================================================
begin;

alter table public.mailing_campaigns
  add column if not exists survey_id uuid references public.surveys(id) on delete set null;

notify pgrst, 'reload schema';
commit;
