-- =============================================================================
-- 120_content_center.sql
-- Content Center becomes its own module; tasks can attach to a post; events can
-- be grouped as recurring occurrences of a series.
-- =============================================================================
begin;

-- Tasks can belong to a social post (e.g. "shoot the reel", "write caption").
alter table public.tasks add column if not exists post_id uuid references public.social_posts(id) on delete set null;
create index if not exists tasks_post_idx on public.tasks (post_id) where post_id is not null;

-- Occurrences generated together as a series point back at the first event.
alter table public.events add column if not exists recurrence_parent_id uuid references public.events(id) on delete set null;
create index if not exists events_recurrence_idx on public.events (recurrence_parent_id) where recurrence_parent_id is not null;

-- Content Center module + Marketing order: Content Center, Social Calendar,
-- Guests, Surveys, Campaigns (Subscribers kept last, unlisted by Keith).
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('content-center', 'Content Center', 'Marketing', 5, true, true, false, array['marketing','full','partners'])
on conflict (key) do update set label = excluded.label, nav_group = 'Marketing', sort_order = 5, built = true, signed_off = true;

update public.module_registry set sort_order = 10 where key = 'social-calendar';
update public.module_registry set sort_order = 20 where key = 'guests';
update public.module_registry set sort_order = 30 where key = 'surveys';
update public.module_registry set sort_order = 40 where key = 'campaigns';
update public.module_registry set sort_order = 50 where key = 'subscribers';

commit;
