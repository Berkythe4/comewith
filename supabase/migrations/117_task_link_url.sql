-- =============================================================================
-- 117_task_link_url.sql
-- A task can carry a URL (e.g. the "Post the episode" production task links to
-- the episode preview page, so opening it from the calendar gives Janelle the
-- tracklist + copy). Nullable; most tasks won't have one.
-- =============================================================================
begin;
alter table public.tasks add column if not exists link_url text;
commit;
