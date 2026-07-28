-- 130: radio episode planning — future skeletons, assignment, recap notes,
-- social-post link, and the DJ share-link foundation.
--
-- 'planned' = a future episode skeleton (assigned, not yet the single active
-- 'building' one). The one-building partial unique index still holds (it only
-- constrains status='building'), so many 'planned' skeletons can coexist.
alter table public.sc_playlists drop constraint if exists sc_playlists_status_check;
alter table public.sc_playlists add constraint sc_playlists_status_check
  check (status = any (array['planned','building','testing','live','archived']));

alter table public.sc_playlists
  add column if not exists recap_notes         text,          -- INTERNAL free-form (recap content notes); never public
  add column if not exists assigned_to         uuid references public.profiles(id),   -- a user in the system
  add column if not exists assigned_actor_id   uuid references public.actors(id),      -- a DJ (actor) given temp access
  add column if not exists dj_token            text,          -- shareable scoped-access token for the assigned DJ
  add column if not exists dj_search_params    jsonb;         -- pre-set filters the DJ's scoped view is locked to
create unique index if not exists uq_sc_playlists_dj_token on public.sc_playlists (dj_token) where dj_token is not null;

-- Tag a radio episode to a social post (so Janelle sees/accesses it from the calendar).
alter table public.social_posts add column if not exists station_id uuid references public.sc_playlists(id);
create index if not exists idx_social_posts_station on public.social_posts (station_id) where station_id is not null;
