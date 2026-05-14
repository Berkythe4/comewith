-- =============================================================================
-- 002_profiles.sql
-- The profiles table extends Supabase's built-in auth.users table with
-- application-specific fields (role, name, must_change_password).
-- One row per auth.users row, linked by id.
-- =============================================================================

create table public.profiles (
  id              uuid primary key references auth.users(id) on delete cascade,
  email           text not null unique,
  full_name       text,
  role            text not null default 'customer'
                    check (role in ('master_admin', 'sub_admin', 'customer')),
  must_change_password boolean not null default false,
  phone           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_profiles_email on public.profiles(email) where deleted_at is null;
create index idx_profiles_role  on public.profiles(role)  where deleted_at is null;

create trigger set_updated_at
  before update on public.profiles
  for each row execute function public.handle_updated_at();

-- =============================================================================
-- Auto-create profile when a new auth.users row is inserted (signup or invite).
-- Defaults role to 'customer'. Berky's row is upgraded manually post-migration.
-- =============================================================================
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, full_name)
  values (
    new.id,
    new.email,
    coalesce(new.raw_user_meta_data->>'full_name', '')
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- =============================================================================
-- RLS policies for profiles
-- =============================================================================
alter table public.profiles enable row level security;

-- Users can read their own profile.
create policy "Users can read own profile"
  on public.profiles for select
  using (auth.uid() = id);

-- Admins can read all profiles.
create policy "Admins can read all profiles"
  on public.profiles for select
  using (public.is_admin());

-- Users can update their own profile (but cannot change role).
create policy "Users can update own profile"
  on public.profiles for update
  using (auth.uid() = id)
  with check (
    auth.uid() = id
    and role = (select role from public.profiles where id = auth.uid())
  );

-- Only master_admin can change roles or create admin profiles.
create policy "Master admin can manage all profiles"
  on public.profiles for all
  using (public.is_master_admin());
