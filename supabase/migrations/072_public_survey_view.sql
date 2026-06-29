-- =============================================================================
-- 072_public_survey_view.sql
-- Anon-readable lookup so a public page (e.g. the impact report) can find an
-- event's OPEN, anonymous-allowed survey and link to it. Exposes only the public
-- token (which is meant to be public for an open anonymous survey) — never invites,
-- responses, or draft/closed surveys.
-- =============================================================================
begin;

create or replace view public.v_public_survey as
  select s.id, s.event_id, s.public_token, s.title
  from public.surveys s
  where s.status = 'open' and s.allow_anonymous = true;

grant select on public.v_public_survey to anon, authenticated;

notify pgrst, 'reload schema';
commit;
