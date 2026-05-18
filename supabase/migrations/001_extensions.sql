-- =============================================================================
-- 001_extensions.sql
-- Enable required Postgres extensions and create helper functions.
-- Run this FIRST on both staging and production projects.
-- =============================================================================

begin;

-- Allow function bodies to reference tables created in later migrations.
-- The helper functions below reference public.profiles, which is created in 002.
set local check_function_bodies = off;

-- UUID generation
create extension if not exists "pgcrypto";

-- Cron jobs (used in Phase 9 for automation)
create extension if not exists "pg_cron";

-- HTTP requests from Postgres (used by pg_cron to call Edge Functions)
create extension if not exists "pg_net";

-- =============================================================================
-- Helper functions
-- =============================================================================

-- Returns the current authenticated user's role from the profiles table.
-- Used in RLS policies throughout the schema.
create or replace function public.current_user_role()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select role from public.profiles where id = auth.uid()
$$;

-- Returns true if current user is master_admin or sub_admin.
create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    (select role in ('master_admin', 'sub_admin') from public.profiles where id = auth.uid()),
    false
  )
$$;

-- Returns true only if current user is master_admin (top-level operations).
create or replace function public.is_master_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    (select role = 'master_admin' from public.profiles where id = auth.uid()),
    false
  )
$$;

-- Auto-update updated_at column on row UPDATE.
-- Attach to any table with: create trigger set_updated_at before update on <table>
--   for each row execute function public.handle_updated_at();
create or replace function public.handle_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

-- Soft delete helper: sets deleted_at = now() instead of true DELETE.
create or replace function public.handle_soft_delete()
returns trigger
language plpgsql
as $$
begin
  new.deleted_at = now();
  return new;
end;
$$;

commit;
