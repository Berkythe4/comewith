-- =============================================================================
-- 041_staff_access_scaffold.sql
-- Staff access model — ADDITIVE SCAFFOLDING ONLY. Safe to apply to prod.
--
-- This migration changes NO existing table policy and removes NO existing access.
-- It only adds:
--   1. profiles.staff_role               (operations | marketing | full)
--   2. backfill existing sub_admins -> 'full' so the nav gate never empties them
--   3. module_registry                   (one row per dashboard module/tab)
--   4. user_module_access                (per-person grant/revoke overrides)
--   5. public.user_can_access_module()   (the effective-access function)
--
-- The actual hard RLS enforcement that USES user_can_access_module() lives in
-- 042 (authored, NOT yet applied — it needs the Events-hub cross-table coupling
-- reviewed first). The dashboard's client-side nav gate uses the same logic in
-- JS and is live immediately. Financial gating lives in 043 (author-only).
--
-- Access axes (orthogonal, per Keith):
--   * staff_role  = WORKFLOW SCOPE  (which module groups a person works in)
--   * signed_off  = RELEASE GATE    (Keith approves a module -> team can see it)
-- Effective (non-master): module is signed_off AND not master_only
--                         AND (role default-grants it OR a per-user grant),
--                         AND no per-user revoke. master_admin always sees all.
-- =============================================================================
begin;

-- 1. staff_role on profiles ---------------------------------------------------
alter table public.profiles
  add column if not exists staff_role text
    check (staff_role in ('operations', 'marketing', 'full'));

comment on column public.profiles.staff_role is
  'Workflow scope for non-master admins: operations | marketing | full. '
  'NULL for master_admin (sees everything) and customers (no dashboard).';

-- 2. Keep existing sub_admins fully scoped so the new nav gate never strands
--    them. (Today: liz@comewith.org.) master_admin is intentionally left NULL.
update public.profiles
  set staff_role = 'full'
  where role = 'sub_admin' and staff_role is null;

-- 3. module_registry ----------------------------------------------------------
-- key       = the dashboard data-tab value (the menu-to-panel contract key)
-- nav_group = section header in the sidebar
-- built     = the module exists and pulls data (false only for placeholders)
-- signed_off= Keith has done his round-2 review and released it to staff
-- master_only = never visible to any staff role (financials, rollups, team mgmt)
-- default_roles = staff_roles that get the module by default (before overrides)
create table if not exists public.module_registry (
  key            text primary key,
  label          text not null,
  nav_group      text not null,
  sort_order     int  not null,
  built          boolean not null default true,
  signed_off     boolean not null default false,
  signed_off_at  timestamptz,
  master_only    boolean not null default false,
  default_roles  text[] not null default '{}',
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now()
);

create trigger set_updated_at
  before update on public.module_registry
  for each row execute function public.handle_updated_at();

alter table public.module_registry enable row level security;

-- Any signed-in admin may READ the registry (the nav needs it); only master may write.
create policy "Admins can read module registry"
  on public.module_registry for select
  using (public.is_admin());
create policy "Master admin manages module registry"
  on public.module_registry for all
  using (public.is_master_admin())
  with check (public.is_master_admin());

-- Seed. Events is the only signed-off module today. Social Calendar is the only
-- unbuilt one. Finance + Strategy + Team are master_only. 'full' staff_role
-- appears in every non-master-only module (== "everything except financials").
insert into public.module_registry
  (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values
  ('inquiries',       'Inquiries',       'Sales',      10, true,  false, false, '{operations,full}'),
  ('agreements',      'Agreements',      'Sales',      20, true,  false, false, '{operations,full}'),
  ('clients',         'Clients',         'Sales',      30, true,  false, false, '{operations,full}'),
  ('events',          'Events',          'Operations', 40, true,  true,  false, '{operations,full}'),
  ('venues',          'Venues',          'Operations', 50, true,  false, false, '{operations,full}'),
  ('equipment',       'Equipment',       'Operations', 60, true,  false, false, '{operations,full}'),
  ('templates',       'Templates',       'Operations', 70, true,  false, false, '{operations,full}'),
  ('income',          'Income',          'Finance',    80, true,  false, true,  '{}'),
  ('expenses',        'Expenses',        'Finance',    90, true,  false, true,  '{}'),
  ('sponsors',        'Sponsors',        'Partners',  100, true,  false, false, '{operations,full}'),
  ('sponsorships',    'Sponsorships',    'Partners',  110, true,  false, false, '{operations,full}'),
  ('artists',         'Artists',         'Partners',  120, true,  false, false, '{operations,full}'),
  ('guests',          'Guests',          'Audience',  130, true,  false, false, '{marketing,full}'),
  ('subscribers',     'Subscribers',     'Audience',  140, true,  false, false, '{marketing,full}'),
  ('campaigns',       'Campaigns',       'Audience',  150, true,  false, false, '{marketing,full}'),
  ('social-calendar', 'Social Calendar', 'Audience',  160, false, false, false, '{marketing,full}'),
  ('strategy',        'Strategy',        'Insights',  170, true,  false, true,  '{}'),
  ('notes',           'Notes',           'Insights',  180, true,  false, false, '{operations,marketing,full}'),
  ('team',            'Team',            'Insights',  190, true,  true,  true,  '{}')
on conflict (key) do nothing;

-- Mark the signed-off timestamp for anything seeded as already approved.
update public.module_registry set signed_off_at = now()
  where signed_off and signed_off_at is null;

-- 4. user_module_access -- per-person overrides on top of the role default -----
create table if not exists public.user_module_access (
  user_id     uuid not null references public.profiles(id) on delete cascade,
  module_key  text not null references public.module_registry(key) on delete cascade,
  access      text not null check (access in ('grant', 'revoke')),
  created_at  timestamptz not null default now(),
  created_by  uuid references public.profiles(id),
  primary key (user_id, module_key)
);

alter table public.user_module_access enable row level security;

-- A user may read their own overrides (the nav needs them); master manages all.
create policy "Users read own module overrides"
  on public.user_module_access for select
  using (auth.uid() = user_id or public.is_master_admin());
create policy "Master admin manages module overrides"
  on public.user_module_access for all
  using (public.is_master_admin())
  with check (public.is_master_admin());

-- 5. Effective-access function ------------------------------------------------
-- True if the current user may use module p_key. master_admin: always.
-- Non-master: signed_off AND not master_only AND role-or-override grants it.
-- A per-user override always wins over the role default; signed_off is absolute
-- (even an explicit grant cannot reveal a module Keith has not released).
create or replace function public.user_can_access_module(p_key text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select case
    when public.is_master_admin() then true
    else coalesce((
      select m.signed_off
         and m.built
         and not m.master_only
         and coalesce(
               (select uma.access
                  from public.user_module_access uma
                 where uma.user_id = auth.uid()
                   and uma.module_key = p_key),
               case when (select p.staff_role from public.profiles p where p.id = auth.uid())
                         = any(m.default_roles)
                    then 'grant' else 'revoke' end
             ) = 'grant'
      from public.module_registry m
      where m.key = p_key
    ), false)
  end
$$;

grant execute on function public.user_can_access_module(text) to authenticated;

commit;

-- =============================================================================
-- POST-APPLY VERIFICATION (run as anon + as berky):
--   * select count(*) from module_registry;                 -- expect 19
--   * select key from module_registry where signed_off;     -- expect events, team
--   * select staff_role from profiles where email='liz@comewith.org'; -- 'full'
--   * Financial views still anon-revoked (unchanged by this migration).
-- ROLLBACK (if ever needed):
--   drop function if exists public.user_can_access_module(text);
--   drop table if exists public.user_module_access;
--   drop table if exists public.module_registry;
--   alter table public.profiles drop column if exists staff_role;
-- =============================================================================
