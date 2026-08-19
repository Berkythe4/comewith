-- =============================================================================
-- 150_gear_watch_facebook_source.sql
-- Allow 'facebook' as a gear_watch_hits source.
--
-- Authored and APPLIED as 148, then renumbered to 150: the laptop landed its own
-- 148_applied_migrations.sql on master while this was open — the same collision
-- CLAUDE.md warns about, and the second one in four days. Prod's
-- applied_migrations row was corrected from 148 to 150 by hand at the same time,
-- and 148 backfilled to the file that actually owns that number. Duplicate
-- numbers conflict in PROD, not in git: nothing here failed loudly.
--
-- 146 wrote `check (source in ('reverb','ebay','craigslist','manual'))`. Adding
-- the Facebook Marketplace source without this would fetch, score and then throw
-- every hit away on insert — and because each insert is checked individually and
-- the error is swallowed per row, the scan would report a cheerful "ok" while
-- silently storing nothing. Exactly the failure mode §24 exists to prevent.
--
-- Facebook Marketplace has no public API from Meta (verified 2026-08-19), so the
-- source runs through Apify and is billed per result. It is off unless
-- APIFY_TOKEN is set, and reports NOT CONFIGURED rather than zero results.
-- =============================================================================
begin;

alter table public.gear_watch_hits
  drop constraint if exists gear_watch_hits_source_check;

alter table public.gear_watch_hits
  add constraint gear_watch_hits_source_check
  check (source in ('reverb', 'ebay', 'craigslist', 'facebook', 'manual'));

notify pgrst, 'reload schema';
commit;

-- DOWN:
--   alter table public.gear_watch_hits drop constraint gear_watch_hits_source_check;
--   alter table public.gear_watch_hits add constraint gear_watch_hits_source_check
--     check (source in ('reverb','ebay','craigslist','manual'));
