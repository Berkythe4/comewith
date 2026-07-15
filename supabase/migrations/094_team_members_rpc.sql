-- =============================================================================
-- 094_team_members_rpc.sql
-- Fixes Supabase security advisor "auth_users_exposed" (advisory email
-- 2026-07-12): v_team_members is an API-exposed view that reads auth.users.
-- Rows are gated by WHERE is_master_admin() and anon is revoked, but the view
-- runs as owner and is granted to `authenticated`, so the linter — and
-- defense-in-depth — flag it: any future bug in the gate would expose
-- auth.users through PostgREST.
-- Fix: drop the view (removes auth.users from the API-exposed surface) and
-- replace it with get_team_members(), a SECURITY DEFINER RPC with the same
-- master-only gate (non-masters get 0 rows). dashboard.html switches
-- sb.from('v_team_members') -> sb.rpc('get_team_members').
-- =============================================================================
begin;

drop view if exists public.v_team_members;

create or replace function public.get_team_members()
returns table (
  id uuid,
  email text,
  full_name text,
  role text,
  staff_role text,
  phone text,
  must_change_password boolean,
  created_at timestamptz,
  deleted_at timestamptz,
  last_sign_in_at timestamptz,
  joined_at timestamptz,
  email_confirmed boolean
)
language sql
stable
security definer
set search_path = public
as $$
  select
    p.id, p.email, p.full_name, p.role, p.staff_role, p.phone,
    p.must_change_password, p.created_at, p.deleted_at,
    u.last_sign_in_at,
    u.created_at as joined_at,
    (u.email_confirmed_at is not null) as email_confirmed
  from public.profiles p
  join auth.users u on u.id = p.id
  where public.is_master_admin();
$$;

-- Functions are EXECUTE-granted to PUBLIC by default — revoke, then grant
-- only the roles that need it.
revoke all on function public.get_team_members() from public, anon;
grant execute on function public.get_team_members() to authenticated, service_role;

commit;
-- POST-APPLY CHECKS:
--   * anon REST POST /rest/v1/rpc/get_team_members -> 401/permission denied
--   * non-master authenticated -> 0 rows; master -> all rows
--   * no public view reads auth.users -> advisor lint auth_users_exposed clears
