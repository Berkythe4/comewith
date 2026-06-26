-- 065_artist_profiles.sql
-- Public artist profiles: bio, photo, socials, a "show on collective" flag, an
-- ordering rank, and a self-edit token. Plus public views for the collective,
-- an artist's gigs, and content tagged to them (recap videos carry an optional
-- artist_id per item).

alter table public.actors
  add column if not exists bio text,
  add column if not exists photo_path text,
  add column if not exists soundcloud text,
  add column if not exists tiktok text,
  add column if not exists public_profile boolean not null default false,
  add column if not exists collective_rank int not null default 100,
  add column if not exists edit_token uuid not null default gen_random_uuid();

-- Seed: surface existing DJs / artists on the collective so the section isn't
-- empty (each can be toggled off individually later).
update public.actors a
set public_profile = true
where a.deleted_at is null
  and exists (select 1 from public.actor_roles r
              where r.actor_id = a.id and r.active and r.role in ('dj','artist'));

-- Collective (public-facing roster)
create or replace view public.v_public_artists as
  select a.id, a.display_name, a.photo_path, a.bio,
         a.instagram, a.soundcloud, a.tiktok, a.website, a.collective_rank
  from public.actors a
  where a.public_profile = true and a.deleted_at is null
  order by a.collective_rank asc, a.display_name asc;

-- Gigs an artist has played (public + completed events only — no leaking
-- unannounced future bookings)
create or replace view public.v_artist_gigs as
  select ep.actor_id, e.id as event_id, e.name as event_name, e.event_date,
         v.name as venue_name, ep.role
  from public.event_participants ep
  join public.events e on e.id = ep.event_id and e.deleted_at is null
  left join public.venues v on v.id = e.venue_id
  where e.is_public = true or e.status = 'completed';

-- Content tagged to an artist (recap videos across events carry artist_id)
create or replace view public.v_artist_content as
  select (item->>'artist_id')::uuid as actor_id,
         item->>'url'   as url,
         item->>'label' as label,
         e.id as event_id, e.name as event_name, e.event_date
  from public.events e,
       lateral jsonb_array_elements(coalesce(e.recap_videos, '[]'::jsonb)) as item
  where e.deleted_at is null
    and coalesce(item->>'artist_id','') <> '';

grant select on public.v_public_artists, public.v_artist_gigs, public.v_artist_content to anon, authenticated;

notify pgrst, 'reload schema';
