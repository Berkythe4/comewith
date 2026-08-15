-- 140_site_owner.sql
--
-- Martin and Henry get full master_admin. The one thing they must NOT be able to
-- do is unseat Keith as overall site owner.
--
-- Today `master_admin` is the top of the app: the policy "Master admin can manage
-- all profiles" is `for all using (is_master_admin())` with no WITH CHECK, so any
-- master_admin can PATCH any profiles row — including Keith's. Promoting the two
-- of them without a guard would hand each of them the ability to remove the other
-- two, Keith included.
--
-- So: an explicit owner flag, and a trigger that protects that one row.
--
-- Note the demotion vector is NOT only `role`. Setting `deleted_at` on the owner
-- is just as fatal: is_admin() / is_master_admin() / user_can_access_module() all
-- treat a deactivated profile as no-role (the 098 deactivation contract), so a
-- single UPDATE would lock Keith out completely while leaving his role reading
-- 'master_admin'. The guard covers role, deleted_at, is_owner and DELETE together.

-- ── 1. The owner flag ───────────────────────────────────────────────────────
alter table public.profiles
  add column if not exists is_owner boolean not null default false;

comment on column public.profiles.is_owner is
  'Overall site owner. Exactly one profile carries this. Guarded by protect_site_owner() — only the owner can change the owner row''s role/deleted_at/is_owner, or hand ownership to someone else.';

update public.profiles
   set is_owner = true
 where id = 'bd71f88d-fd0a-4e3a-a712-7ac958318c8b';   -- berky@comewith.org (Keith)

-- ── 2. The guard ────────────────────────────────────────────────────────────
create or replace function public.protect_site_owner()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  actor uuid := auth.uid();
begin
  -- No JWT = service role, a Management-API/SQL session, or an Edge Function.
  -- Deliberately NOT restricted: this is Keith's break-glass path, and locking it
  -- would mean a bad row could only be fixed by Supabase support. It also means
  -- the guard is an APP-level control, not a defence against anyone holding the
  -- service-role key or a Supabase project token. Those stay Keith-only.
  if actor is null then
    return case when tg_op = 'DELETE' then old else new end;
  end if;

  if tg_op = 'INSERT' then
    if new.is_owner then
      raise exception 'Ownership cannot be granted at insert time'
        using errcode = '42501';
    end if;
    return new;
  end if;

  if tg_op = 'DELETE' then
    if old.is_owner and actor <> old.id then
      raise exception 'The site owner''s account cannot be deleted by another user'
        using errcode = '42501';
    end if;
    return old;
  end if;

  -- UPDATE. Everything else on the owner's row (name, phone, staff_role) stays
  -- editable by an admin — it's the three fields that decide whether Keith still
  -- runs this place that are locked.
  if old.is_owner and actor <> old.id then
    if new.role is distinct from old.role
       or new.deleted_at is distinct from old.deleted_at
       or new.is_owner is distinct from old.is_owner then
      raise exception 'Only the site owner can change the site owner''s role, active status or ownership'
        using errcode = '42501';
    end if;
  end if;

  -- Ownership is transferred BY the owner, never taken.
  if new.is_owner and not old.is_owner then
    if not exists (select 1 from public.profiles p where p.id = actor and p.is_owner) then
      raise exception 'Only the current site owner can transfer ownership'
        using errcode = '42501';
    end if;
  end if;

  return new;
end;
$$;

drop trigger if exists protect_site_owner on public.profiles;
create trigger protect_site_owner
  before insert or update or delete on public.profiles
  for each row execute function public.protect_site_owner();

-- Convenience for the UI: "am I the owner?" without exposing other rows.
create or replace function public.is_site_owner()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    (select is_owner and deleted_at is null from public.profiles where id = auth.uid()),
    false
  )
$$;

grant execute on function public.is_site_owner() to authenticated;

-- ── 3. Surface it in the Users tab ──────────────────────────────────────────
-- With three master_admins, three identical "master" chips tell you nothing about
-- who actually runs the place. get_team_members() is the master-only RPC the Users
-- tab reads (094 replaced the v_team_members view with it after the
-- auth_users_exposed advisory — never expose auth.users through a view). Adding a
-- column to a RETURNS TABLE needs a drop, not a replace.
drop function if exists public.get_team_members();
create function public.get_team_members()
returns table (
  id uuid, email text, full_name text, role text, staff_role text, phone text,
  must_change_password boolean, created_at timestamptz, deleted_at timestamptz,
  last_sign_in_at timestamptz, joined_at timestamptz, email_confirmed boolean,
  is_owner boolean
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
    (u.email_confirmed_at is not null) as email_confirmed,
    p.is_owner
  from public.profiles p
  join auth.users u on u.id = p.id
  where public.is_master_admin();
$$;

grant execute on function public.get_team_members() to authenticated;

-- ── 4. The promotions ───────────────────────────────────────────────────────
update public.profiles
   set role = 'master_admin'
 where email in ('martin@comewith.org', 'henry@comewith.org')
   and deleted_at is null;
