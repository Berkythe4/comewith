-- 098: User deactivation enforcement + Site Editor/Review into Team HQ
--
-- Deactivating a user = setting profiles.deleted_at (master-only, via the
-- Users tab). For that to actually revoke access — not just hide the row —
-- the role helpers must treat a deactivated profile as no-role: is_admin() /
-- is_master_admin() gate nearly every RLS policy, and user_can_access_module()
-- has a staff branch that never consults is_admin(), so it gets its own guard.
-- A deactivated user can still authenticate, but every policy then denies them.

create or replace function public.is_admin()
returns boolean language sql stable security definer set search_path = public as $$
  select coalesce(
    (select role in ('master_admin', 'sub_admin') and deleted_at is null
       from public.profiles where id = auth.uid()),
    false
  )
$$;

create or replace function public.is_master_admin()
returns boolean language sql stable security definer set search_path = public as $$
  select coalesce(
    (select role = 'master_admin' and deleted_at is null
       from public.profiles where id = auth.uid()),
    false
  )
$$;

create or replace function public.user_can_access_module(p_key text)
returns boolean language sql stable security definer set search_path = public as $$
  select case
    when public.is_master_admin() then true
    when not public.is_admin() then false   -- 098: deactivated staff get nothing
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

-- ---- nav: Site Editor + Site Review move from Marketing to Team HQ ---------
update public.module_registry set nav_group = 'Team HQ', sort_order = 50 where key = 'site-editor';
update public.module_registry set nav_group = 'Team HQ', sort_order = 60 where key = 'site-review';
