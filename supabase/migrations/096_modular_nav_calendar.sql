-- 096: Modular product-nav regroup + master Calendar module + task milestones
--
-- Regroups module_registry into sellable product modules (Keith, 2026-07-17):
--   Pinned    : calendar (new — Master Calendar & Tasks, pinned above the groups)
--   Workflow  : events, inquiries, pricing, agreements, clients, templates, vendors
--   Finance   : income, expenses, sponsors, sponsorships
--   Marketing : campaigns, subscribers, surveys, social-calendar, guests,
--               site-editor, site-review
--   Venues    : venues, equipment
--   Artists   : artists, actors (relabeled "People & Orgs")
--   Radio     : ra-market (Artist Radio), market (Best Nights)
--   Team HQ   : team (Users), conversations, strategy, notes
-- The dashboard renders 'Pinned' items flat at the top and the rest as
-- collapsible sections (NAV_GROUP_ORDER in dashboard.html must match).
-- New tables inherit grants from 013 default privileges; no anon grants here.

-- ---- new Calendar module ---------------------------------------------------
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('calendar', 'Calendar & Tasks', 'Pinned', 1, true, true, false, '{operations,marketing,full}')
on conflict (key) do update
  set label = excluded.label, nav_group = excluded.nav_group, sort_order = excluded.sort_order;

-- ---- regroup existing modules ---------------------------------------------
update public.module_registry set nav_group = 'Workflow',  sort_order = 10  where key = 'events';
update public.module_registry set nav_group = 'Workflow',  sort_order = 20  where key = 'inquiries';
update public.module_registry set nav_group = 'Workflow',  sort_order = 30  where key = 'pricing';
update public.module_registry set nav_group = 'Workflow',  sort_order = 40  where key = 'agreements';
update public.module_registry set nav_group = 'Workflow',  sort_order = 50  where key = 'clients';
update public.module_registry set nav_group = 'Workflow',  sort_order = 60  where key = 'templates';
update public.module_registry set nav_group = 'Workflow',  sort_order = 70  where key = 'vendors';

update public.module_registry set nav_group = 'Finance',   sort_order = 10  where key = 'income';
update public.module_registry set nav_group = 'Finance',   sort_order = 20  where key = 'expenses';
update public.module_registry set nav_group = 'Finance',   sort_order = 30  where key = 'sponsors';
update public.module_registry set nav_group = 'Finance',   sort_order = 40  where key = 'sponsorships';

update public.module_registry set nav_group = 'Marketing', sort_order = 10  where key = 'campaigns';
update public.module_registry set nav_group = 'Marketing', sort_order = 20  where key = 'subscribers';
update public.module_registry set nav_group = 'Marketing', sort_order = 30  where key = 'surveys';
update public.module_registry set nav_group = 'Marketing', sort_order = 40  where key = 'social-calendar';
update public.module_registry set nav_group = 'Marketing', sort_order = 50  where key = 'guests';
update public.module_registry set nav_group = 'Marketing', sort_order = 60  where key = 'site-editor';
update public.module_registry set nav_group = 'Marketing', sort_order = 70  where key = 'site-review';

update public.module_registry set nav_group = 'Venues',    sort_order = 10  where key = 'venues';
update public.module_registry set nav_group = 'Venues',    sort_order = 20  where key = 'equipment';

update public.module_registry set nav_group = 'Artists',   sort_order = 10  where key = 'artists';
update public.module_registry set nav_group = 'Artists',   sort_order = 20, label = 'People & Orgs' where key = 'actors';

update public.module_registry set nav_group = 'Radio',     sort_order = 10  where key = 'ra-market';
update public.module_registry set nav_group = 'Radio',     sort_order = 20  where key = 'market';

update public.module_registry set nav_group = 'Team HQ',   sort_order = 10  where key = 'team';
update public.module_registry set nav_group = 'Team HQ',   sort_order = 20  where key = 'conversations';
update public.module_registry set nav_group = 'Team HQ',   sort_order = 30  where key = 'strategy';
update public.module_registry set nav_group = 'Team HQ',   sort_order = 40  where key = 'notes';

-- ---- task milestones -------------------------------------------------------
-- Milestones are the only tasks that render on the master calendar grid;
-- ordinary tasks live on the board below it.
alter table public.tasks add column if not exists milestone boolean not null default false;
create index if not exists idx_tasks_due on public.tasks (due_date) where deleted_at is null;
