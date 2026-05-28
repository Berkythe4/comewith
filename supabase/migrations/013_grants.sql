-- =============================================================================
-- 013_grants.sql
-- Standard Supabase grant pattern for the public schema.
--
-- Phase 0 oversight: migrations 002-012 created tables with RLS policies but
-- never granted the anon / authenticated / service_role roles base table-level
-- privileges. Supabase Cloud projects normally provide these defaults, but if
-- the public schema is ever dropped and recreated (e.g. the rollback in
-- PHASE0_README), the defaults are wiped. Without these grants, even a
-- "with check (true)" INSERT policy rejects rows with a 42501 RLS violation,
-- because the role can't reach the table to evaluate policies in the first
-- place.
--
-- Row-level filtering is still enforced by the RLS policies on each table;
-- these grants only let the roles get to the policy evaluation step.
--
-- Safe to re-run: GRANT and ALTER DEFAULT PRIVILEGES are idempotent.
-- =============================================================================

grant usage on schema public to anon, authenticated, service_role;

grant all on all tables    in schema public to anon, authenticated, service_role;
grant all on all sequences in schema public to anon, authenticated, service_role;
grant all on all functions in schema public to anon, authenticated, service_role;

-- Future tables created in this schema inherit the same grants
alter default privileges in schema public
  grant all on tables    to anon, authenticated, service_role;
alter default privileges in schema public
  grant all on sequences to anon, authenticated, service_role;
alter default privileges in schema public
  grant all on functions to anon, authenticated, service_role;
