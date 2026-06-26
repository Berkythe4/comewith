-- =============================================================================
-- 062_site_content.sql
-- Editable public-site content (a tiny CMS). Every editable text/image/embed on
-- the site is a key in site_content; the public site reads it (anon) and falls
-- back to inline defaults; the dashboard "Site Editor" writes it. Images (logo,
-- hero) live in the public 'event-photos' bucket under a site/ prefix.
-- =============================================================================
begin;
create table if not exists public.site_content (
  key text primary key,
  value text,
  updated_at timestamptz not null default now(),
  updated_by uuid references public.profiles(id)
);
alter table public.site_content enable row level security;
-- public site reads everything; only the site-editor module (or master) writes.
create policy "Site content public read" on public.site_content for select using (true);
create policy "Site content edit" on public.site_content for all
  using (public.user_can_access_module('site-editor') or public.is_master_admin())
  with check (public.user_can_access_module('site-editor') or public.is_master_admin());
grant select on public.site_content to anon, authenticated;

insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('site-editor', 'Site Editor', 'Insights',
        (select coalesce(max(sort_order),0)+1 from public.module_registry),
        true, true, false, array['marketing','full'])
on conflict (key) do update set built = true, signed_off = true, label = 'Site Editor';
commit;
