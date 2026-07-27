-- =============================================================================
-- 122_event_content_confirmed.sql
-- Let a content stage be CONFIRMED as "intentionally none" on the pipeline, so
-- an event with no photos/video/recap-by-design reads as resolved (a calm
-- indigo "complete") instead of a nagging amber "pending". Stored per stage.
-- =============================================================================
begin;
alter table public.events add column if not exists content_confirmed jsonb not null default '{}'::jsonb;
commit;
