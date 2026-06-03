-- =============================================================================
-- 029_sponsorship_actor_and_donor_role.sql  (additive, reversible)
-- Enables loading DI data into the actor model:
--  - sponsorships.sponsor_id -> nullable, so a sponsorship can attach to an
--    actor (actor_id, added in 023) without a legacy sponsors row.
--  - actor_roles.role -> add 'donor' (Keith donates out-of-pocket to his own
--    events; counted as raised, attributed to Keith-the-actor).
-- APPLIED TO PROD 2026-06-02 via Management API.
-- DOWN:
--   alter table public.sponsorships alter column sponsor_id set not null;  -- only if no null rows
--   alter table public.actor_roles drop constraint actor_roles_role_check;
--   alter table public.actor_roles add constraint actor_roles_role_check check (role in
--     ('artist','dj','contractor','customer','sponsor','team','performer','painter','dancer','vendor','venue_contact','host','crew'));
-- =============================================================================
alter table public.sponsorships alter column sponsor_id drop not null;
alter table public.actor_roles drop constraint if exists actor_roles_role_check;
alter table public.actor_roles add constraint actor_roles_role_check check (role in
 ('artist','dj','contractor','customer','sponsor','team','performer','painter','dancer','vendor','venue_contact','host','crew','donor'));
