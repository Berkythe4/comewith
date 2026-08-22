-- ============================================================
-- COME WITH — 202 lock down the 197 trigger guards' EXECUTE
--
-- Postgres grants EXECUTE on a new function to PUBLIC, and `anon` inherits it.
-- 197 created plan_frozen_guard() and plan_version_guard() without saying
-- otherwise, so both ended up anon-executable (`=X/postgres` in proacl).
--
-- HONEST SCOPE: this is hygiene, not a hole. Both are `security invoker` TRIGGER
-- functions — Postgres refuses a direct call ("trigger functions can only be
-- called as triggers") before a line of the body runs, and invoker means they
-- would carry no extra privilege even if they did. Nothing was exposed.
--
-- It is closed anyway because the standing rule (LEARNINGS §45) is that a
-- function's grants are stated rather than inherited: `revoke ... from anon`
-- alone is a silent no-op, and the only reliable form is
-- `revoke all ... from public, anon` followed by an explicit grant. Leaving one
-- function on the default is how the next one that DOES matter gets missed.
-- ============================================================
begin;

revoke all on function public.plan_frozen_guard()  from public, anon;
revoke all on function public.plan_version_guard() from public, anon;

-- The triggers themselves run as the table owner and do not need these grants;
-- they are here so an admin can still inspect/call them while debugging.
grant execute on function public.plan_frozen_guard()  to authenticated, service_role;
grant execute on function public.plan_version_guard() to authenticated, service_role;

commit;

-- DOWN:
--   grant execute on function public.plan_frozen_guard()  to public;
--   grant execute on function public.plan_version_guard() to public;
