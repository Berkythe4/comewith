-- Post-apply invariant check for prod.
--
--   SBP_REF=yaytdosxfhcqatmhctzk python db.py supabase/checks/post_apply.sql
--
-- Deliberately ONE statement so the whole verification costs a single
-- Management API call (and a single approval) instead of six.
--
-- Every row must read PASS. INFO rows are for eyeballing, not gating.
-- This does NOT replace the anon REST GET on the financial views (expect 401)
-- — that tests PostgREST end to end; this tests the grants behind it.
--
-- Two of these checks cried wolf on their first real run (2026-08-15, applying
-- 141). Both were bugs in the check, not in prod, and both failed the same way:
-- they tested for a PROXY of the invariant rather than the invariant. That is
-- worse than not checking at all — a gate that always FAILs teaches you to wave
-- it through, and this file is what every future migration is verified with.
-- What changed is recorded at each check below.

with fin(name) as (
  values ('v_event_summary'), ('v_kpi_event_financials'), ('v_kpi_parties'),
         ('v_kpi_dance_infusion'), ('v_kpi_dashboard')
)

-- E1: the five financial views must not be SELECT-able by anon. That, and only
-- that, is the invariant — it is what the 016/017 regression broke and what 019
-- restored, and it is what the anon REST GET proves with a 401.
--
-- WAS: counted anon grants of ANY privilege_type. The old blanket
-- `grant all on all tables in schema public to anon` left INSERT/UPDATE/DELETE/
-- TRUNCATE/REFERENCES/TRIGGER behind on all five views even after SELECT was
-- revoked, so this FAILed permanently while anon could not read a single row.
-- It could not tell inert residue from a live re-exposure — the one distinction
-- it exists to make.
select
  'financial_views_anon_select' as check_name,
  case when count(*) = 0 then 'PASS' else 'FAIL' end as status,
  coalesce(string_agg(table_name, ', '), 'no anon SELECT - correct') as detail
from information_schema.role_table_grants
where grantee = 'anon'
  and table_schema = 'public'
  and privilege_type = 'SELECT'
  and table_name in (select name from fin)

union all

-- The residue above is inert ONLY because none of the five views is
-- auto-updatable — Postgres refuses a write to a non-auto-updatable view no
-- matter what has been granted. That is an assumption worth testing rather than
-- asserting in a comment: if one of these is ever rewritten into something
-- simple enough to be auto-updatable, the leftover grants stop being harmless.
select
  'financial_views_not_updatable',
  case when count(*) = 0 then 'PASS' else 'FAIL' end,
  coalesce(string_agg(table_name || ' [ins=' || is_insertable_into
                      || ' upd=' || is_updatable || ']', ', '),
           'none auto-updatable - correct')
from information_schema.views
where table_schema = 'public'
  and table_name in (select name from fin)
  and (is_insertable_into = 'YES' or is_updatable = 'YES')

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
--
-- WAS: grepped every helper's body for the literal 'deleted_at'. But the
-- contract can be honoured by DELEGATION, and in prod it is —
-- user_can_access_module carries no deleted_at of its own, it short-circuits on
--   when not public.is_admin() then false   -- 098: deactivated staff get nothing
-- and is_admin() / is_master_admin() hold the guard. So the correct
-- implementation was reported as a violation.
--
-- The exemption is deliberately narrow: it applies ONLY to the derived helper.
-- The two base helpers must still carry deleted_at themselves, or delegation
-- would have nothing underneath it to stand on.
select
  'role_helpers_keep_deleted_at',
  case when count(*) = 0 then 'PASS' else 'FAIL' end,
  coalesce(string_agg(p.proname, ', '), 'all guarded - correct')
from pg_proc p
join pg_namespace n on n.oid = p.pronamespace
where n.nspname = 'public'
  and p.proname in ('is_admin', 'is_master_admin', 'user_can_access_module')
  and pg_get_functiondef(p.oid) not ilike '%deleted_at%'
  and not (p.proname = 'user_can_access_module'
           and pg_get_functiondef(p.oid) ilike '%is_admin(%')

union all

-- Visible, not gating: the non-SELECT anon grants the blanket-grant era left on
-- the financial views. Harmless while the check above passes; worth watching, and
-- worth revoking whenever something else is touching these views anyway.
select
  'financial_views_anon_residue',
  'INFO',
  case when count(*) = 0 then 'none'
       else count(*)::text || ' non-SELECT anon grants across '
            || count(distinct table_name)::text || ' of the 5 views' end
from information_schema.role_table_grants
where grantee = 'anon'
  and table_schema = 'public'
  and privilege_type <> 'SELECT'
  and table_name in (select name from fin)

union all

-- The 016/017 regression looked like a sudden jump in this number.
select
  'anon_grant_inventory',
  'INFO',
  count(*)::text || ' anon table/view grants in public'
from information_schema.role_table_grants
where grantee = 'anon' and table_schema = 'public'

order by 1;
