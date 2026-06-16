-- =============================================================================
-- 039_actor_role_staff.sql  —  add 'staff' to actor_roles.role (additive widening)
-- Same pattern migration 029 used to add 'donor'. Widening a CHECK never invalidates
-- existing rows. Lets event staff/crew from the DI#2 ledger graduate to the actor layer.
-- =============================================================================
alter table public.actor_roles drop constraint if exists actor_roles_role_check;
alter table public.actor_roles add constraint actor_roles_role_check check (role in (
  'artist','dj','contractor','customer','sponsor','team','performer','painter','dancer',
  'vendor','venue_contact','host','crew','donor','staff'));
