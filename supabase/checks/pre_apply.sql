-- Pre-apply introspection for prod.
--
--   1. edit the `targets` list below to the tables/views this migration touches
--   2. SBP_REF=$SBP_REF_PROD python db.py supabase/checks/pre_apply.sql
--
-- Deliberately ONE statement so introspection costs a single Management API
-- call (and a single approval). Run this BEFORE writing the migration, so
-- [VERIFY] refs get reconciled against the live database instead of assumed.
--
-- Empty output for a target means it does not exist yet — which is the answer
-- to "is this additive or a rewrite?"

with targets(name) as (
  values ('link_pages'), ('link_items'), ('v_public_link_pages'), ('v_public_link_items'), ('module_registry')     -- <-- EDIT ME
)

select 'a_object' as kind, c.relname as name,
       case c.relkind
         when 'r' then 'table' when 'v' then 'view'
         when 'm' then 'matview' when 'p' then 'partitioned table'
         else c.relkind::text end
       || case when c.relrowsecurity then ' [RLS on]' else '' end as detail
from pg_class c
join pg_namespace n on n.oid = c.relnamespace
where n.nspname = 'public' and c.relname in (select name from targets)

union all

select 'b_column', table_name,
       column_name || ' ' || data_type
       || case when is_nullable = 'NO' then ' not null' else '' end
       || coalesce(' default ' || column_default, '')
from information_schema.columns
where table_schema = 'public' and table_name in (select name from targets)

union all

select 'c_policy', tablename, policyname || ' [' || cmd || ']'
from pg_policies
where schemaname = 'public' and tablename in (select name from targets)

union all

select 'd_grant', table_name, grantee || ':' || privilege_type
from information_schema.role_table_grants
where table_schema = 'public' and table_name in (select name from targets)

order by 1, 2, 3;
