-- =============================================================================
-- 049_actor_orgs_and_roles.sql  (additive)
--  A) actors.org_id — a person belongs to an org (another actor, kind=org).
--     Powers "assign people to orgs" on the Actors page. (venues.actor_id already
--     exists for "assign venues to orgs".)
--  B) widen actor_roles.role to capture more of the spots a person can fill.
-- Admin-only via existing RLS; no anon grant.
-- =============================================================================
begin;

alter table public.actors add column if not exists org_id uuid references public.actors(id) on delete set null;
create index if not exists idx_actors_org on public.actors(org_id) where org_id is not null;

alter table public.actor_roles drop constraint if exists actor_roles_role_check;
alter table public.actor_roles add constraint actor_roles_role_check check (role in (
  'artist', 'dj', 'contractor', 'customer', 'sponsor', 'team', 'staff', 'performer',
  'painter', 'dancer', 'vendor', 'venue_contact', 'host', 'crew', 'donor',
  'photographer', 'videographer', 'producer', 'organizer', 'volunteer', 'partner', 'designer'));

commit;

-- DOWN: alter table actors drop column org_id; (and restore the prior role check)
