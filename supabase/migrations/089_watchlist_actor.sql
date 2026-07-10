-- =============================================================================
-- 089_watchlist_actor.sql
-- Watchlist artists (esp. Collaborators) can be linked to an actor in the system
-- (or one is created for them), so a watched collaborator ties into the roster.
-- =============================================================================
begin;
alter table public.watchlist add column if not exists actor_id uuid references public.actors(id) on delete set null;
commit;
-- POST: watchlist.actor_id links a watched artist to public.actors.
