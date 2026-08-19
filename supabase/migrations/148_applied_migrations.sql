-- ============================================================
-- COME WITH — 148 applied-migration tracking
--
-- Migrations here are applied as raw SQL through the Management API (db.py),
-- which bypasses the Supabase CLI entirely — so `supabase_migrations.schema_
-- migrations` was never created and prod has NO record of what has been run.
-- With three machines authoring numbers independently, the only way to answer
-- "is 147 live?" has been to introspect objects and infer. That is exactly how
-- 138 got authored twice (see CLAUDE.md).
--
-- This gives it a place to be written down. db.py records each migration file it
-- successfully applies; a dry run (commit swapped for rollback) is deliberately
-- NOT recorded.
--
-- HISTORY BEFORE THIS POINT IS UNKNOWN and is not backfilled — guessing which of
-- 001-147 are live would be worse than an honest gap. Rows accumulate from the
-- first migration applied after this lands. For anything older, introspect.
-- ============================================================
begin;

create table if not exists public.applied_migrations (
  version     text primary key,             -- '147'
  filename    text not null,                -- '147_fpa_pl.sql'
  sha256      text,                         -- content hash, so an edited-then-
                                            -- reapplied file is visible
  applied_at  timestamptz not null default now(),
  applied_by  text,                         -- machine that ran it
  note        text
);

comment on table public.applied_migrations is
  'What has actually been run against this database, written by db.py. Rows only '
  'exist for migrations applied after 148 landed; earlier history is untracked.';

alter table public.applied_migrations enable row level security;

drop policy if exists "Admins can read applied migrations" on public.applied_migrations;
create policy "Admins can read applied migrations"
  on public.applied_migrations for select using (public.is_admin());

-- Never anon-readable: it maps the shape of the schema and who runs it.
revoke all on public.applied_migrations from anon;

commit;

-- DOWN: drop table if exists public.applied_migrations;
