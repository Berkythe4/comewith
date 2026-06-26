-- =============================================================================
-- 058_user_management.sql
-- Backs the revamped user-management tool (master-only):
--   * v_team_members — profiles joined to auth.users for last_sign_in / joined /
--     email-confirmed. Runs as owner (so it can read auth.users) but is gated by
--     a WHERE is_master_admin() so only a master sees any row; revoked from anon.
--   * indexes for the activity feed (audit_log by time/actor) and alias lookups
--     (actors.user_id — a staff login tagged to their performer/DJ actor records).
-- =============================================================================
begin;

create or replace view public.v_team_members as
select
  p.id, p.email, p.full_name, p.role, p.staff_role, p.phone,
  p.must_change_password, p.created_at, p.deleted_at,
  u.last_sign_in_at,
  u.created_at as joined_at,
  (u.email_confirmed_at is not null) as email_confirmed
from public.profiles p
join auth.users u on u.id = p.id
where public.is_master_admin();

revoke all on public.v_team_members from anon;
grant select on public.v_team_members to authenticated;

create index if not exists idx_actors_user      on public.actors(user_id)     where user_id is not null;
create index if not exists idx_audit_occurred    on public.audit_log(occurred_at desc);
create index if not exists idx_audit_actor        on public.audit_log(actor_id);
create index if not exists idx_audit_table        on public.audit_log(table_name);
commit;
-- POST: anon REST GET v_team_members -> 401; a non-master staff -> 0 rows; master -> all.
