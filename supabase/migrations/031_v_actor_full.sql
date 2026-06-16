-- =============================================================================
-- 031_v_actor_full.sql  —  Event Hub sprint (Sprint 1 of the module series)
--
-- Establishes the v_actor_full convention: the canonical "give me an actor with
-- their roles" read. Later module sprints (Artist, Vendor) extend this view to
-- LEFT JOIN their one-to-one actor_<role>_details tables as those tables are
-- added (see docs/ACTOR_DETAILS_PATTERN.md). The actors table stays the universal
-- core; no wide sparse columns, no JSON blob for analyzable fields.
--
-- ADDITIVE ONLY: one new view. No table changes, no data writes, no drops.
-- RLS: created WITH (security_invoker = true) so the view enforces the underlying
-- actors / actor_roles RLS (is_admin()) against EACH caller, instead of the default
-- definer behavior that would bypass RLS (the same definer-bypass that forced the
-- financial views to be anon-revoked). With security_invoker the view is admin-only
-- by construction: admins pass is_admin(); any future external authenticated actor
-- gets zero rows; anon is revoked outright. This is the pattern later actor_*_details
-- joins inherit, so they stay safe as external logins arrive. NEVER blanket-grant
-- anon (013/016/017/019 discipline).
-- =============================================================================
begin;

create or replace view public.v_actor_full
with (security_invoker = true) as
select
  a.*,
  coalesce(
    (select array_agg(ar.role order by ar.role)
       from public.actor_roles ar
      where ar.actor_id = a.id and ar.active = true),
    '{}'
  ) as roles
from public.actors a
where a.deleted_at is null;

comment on view public.v_actor_full is
  'Canonical actor read: every actor with their active roles as a text[]. The base of the actor_*_details pattern (docs/ACTOR_DETAILS_PATTERN.md) — later sprints LEFT JOIN actor_artist_details / actor_vendor_details here. Admin-only via the underlying actor-tables RLS; not exposed to anon.';

-- Least privilege: the view inherits no broad grant. anon gets nothing; admin
-- reads through the authenticated role + is_admin() RLS on actors/actor_roles.
revoke all on public.v_actor_full from anon;

commit;

-- DOWN: drop view if exists public.v_actor_full;
