-- Post-apply invariant check for prod.
--
--   SBP_REF=$SBP_REF_PROD python db.py supabase/checks/post_apply.sql
--
-- Deliberately ONE statement so the whole verification costs a single
-- Management API call (and a single approval) instead of six.
--
-- Every row must read PASS. INFO rows are for eyeballing, not gating.
-- This does NOT replace the anon REST GET on the financial views (expect 401)
-- — that tests PostgREST end to end; this tests the grants behind it.

select
  'financial_views_anon' as check_name,
  case when count(*) = 0 then 'PASS' else 'FAIL' end as status,
  coalesce(string_agg(table_name || ':' || privilege_type, ', '),
           'no anon grants - correct') as detail
from information_schema.role_table_grants
where grantee = 'anon'
  and table_schema = 'public'
  and table_name in (
    'v_event_summary', 'v_kpi_event_financials', 'v_kpi_parties',
    'v_kpi_dance_infusion', 'v_kpi_dashboard')

union all

-- 103: sc_playlists / sc_playlist_tracks carried table-level anon grants from
-- 079. RLS hid the rows, so an anon GET returned 200 [] and looked fine.
select
  'radio_tables_anon',
  case when count(*) = 0 then 'PASS' else 'FAIL' end,
  coalesce(string_agg(distinct table_name, ', '), 'no anon grants - correct')
from information_schema.role_table_grants
where grantee = 'anon'
  and table_schema = 'public'
  and (table_name like 'listener\_%' escape '\'
       or table_name in ('sc_playlists', 'sc_playlist_tracks'))

union all

-- RLS enabled with zero policies denies everyone, including admins, and reads
-- as "secured" in the dashboard.
select
  'rls_enabled_without_policy',
  case when count(*) = 0 then 'PASS' else 'FAIL' end,
  coalesce(string_agg(c.relname, ', '), 'none - correct')
from pg_class c
join pg_namespace n on n.oid = c.relnamespace
where n.nspname = 'public'
  and c.relkind = 'r'
  and c.relrowsecurity
  and not exists (select 1 from pg_policy p where p.polrelid = c.oid)

union all

-- 098 deactivation contract: a role helper that drops the deleted_at guard
-- silently re-grants a deactivated user their role.
select
  'role_helpers_keep_deleted_at',
  case when count(*) = 0 then 'PASS' else 'FAIL' end,
  coalesce(string_agg(p.proname, ', '), 'all guarded - correct')
from pg_proc p
join pg_namespace n on n.oid = p.pronamespace
where n.nspname = 'public'
  and p.proname in ('is_admin', 'is_master_admin', 'user_can_access_module')
  and pg_get_functiondef(p.oid) not ilike '%deleted_at%'

union all

-- The 016/017 regression looked like a sudden jump in this number.
select
  'anon_grant_inventory',
  'INFO',
  count(*)::text || ' anon table/view grants in public'
from information_schema.role_table_grants
where grantee = 'anon' and table_schema = 'public'

order by 1;
