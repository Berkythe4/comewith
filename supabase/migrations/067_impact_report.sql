-- =============================================================================
-- 067_impact_report.sql
-- Admin-editable, per-event PUBLIC impact report (Dance Infusion).
--   * events.impact_report        jsonb   — curated content: hero/narrative text,
--                                            the public audit figures, photo paths.
--   * events.impact_report_public boolean — the PUBLISH toggle (the gate).
--
-- v_public_impact_report is anon-readable and returns a report ONLY when the
-- toggle is on — that single rule is what makes the dashboard toggle "publish".
--
-- Photos live in the existing PUBLIC 'event-photos' bucket; only their storage
-- paths sit inside the jsonb. NO financial views are touched here: the audit
-- figures shown publicly are the CURATED values in the jsonb ("transparency on
-- our terms"), NOT the anon-revoked v_event_summary / v_kpi_* financial views,
-- which stay revoked from anon by design.
-- =============================================================================
begin;

alter table public.events
  add column if not exists impact_report        jsonb   not null default '{}'::jsonb,
  add column if not exists impact_report_public boolean not null default false;

create or replace view public.v_public_impact_report as
  select e.id,
         e.name,
         e.event_date,
         v.name as venue_name,
         e.series,
         e.hero_image_path,
         e.impact_report
  from public.events e
    left join public.venues v on v.id = e.venue_id
  where e.impact_report_public = true
    and e.deleted_at is null
  order by e.event_date desc;

-- Anon read of PUBLIC reports only (the view already filters to public rows).
-- Mirrors the v_public_recap / v_public_events grant pattern. Default privileges
-- from 013 do NOT cover views, so this explicit grant is required and is safe
-- (the view exposes no financial-view data).
grant select on public.v_public_impact_report to anon, authenticated;

notify pgrst, 'reload schema';

commit;
