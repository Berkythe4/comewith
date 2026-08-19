-- ============================================================
-- COME WITH — 174 photos get a photographer, and photographers get a portfolio
--
-- Photos already have a home: the public 'event-photos' bucket plus the
-- event_photos table, uploaded through Event Hub → Photos, which does the
-- HEIC→JPEG conversion, makes a full + thumb pair, and dedupes on file_hash.
-- Nothing about that needs replacing.
--
-- What is missing is WHO TOOK THE PHOTO. Tori Mumtaz was paid $250 for a shoot
-- and there is no way to say which images are hers — so they cannot be pulled
-- into a portfolio for her, and Come With cannot credit her reliably when it
-- publishes one. Both matter: crediting a photographer is usually a condition of
-- using their work, and doing it from memory is how it gets missed.
--
--   photographer_actor_id — who shot it. An actor, so it reuses every existing
--                           person record rather than inventing a name field.
--   credit_line           — an override for the rare case where the credit
--                           differs from the actor's display name (a studio
--                           name, an agreed handle). Null means "use the name".
--   shot_on               — when it was taken, which is not always when it was
--                           uploaded, and is what a portfolio sorts by.
--
-- v_photo_credits gives both directions in one place: what to print under an
-- image, and everything one photographer has shot.
-- ============================================================
begin;

alter table public.event_photos
  add column if not exists photographer_actor_id uuid references public.actors(id),
  add column if not exists credit_line text,
  add column if not exists shot_on date;

comment on column public.event_photos.photographer_actor_id is
  'Who took the photo. Null means unattributed — not the same as "no credit needed".';
comment on column public.event_photos.credit_line is
  'Overrides the photographer''s display name in the printed credit. Null = use the name.';

create index if not exists idx_event_photos_photographer
  on public.event_photos (photographer_actor_id)
  where photographer_actor_id is not null;

-- One row per photo, with the credit already assembled and a public URL, so
-- neither the site nor a portfolio export has to rebuild either.
create or replace view public.v_photo_credits as
select
  p.id                                   as photo_id,
  p.event_id,
  e.name                                 as event_name,
  e.event_date,
  p.storage_path,
  p.thumb_path,
  p.caption,
  p.is_public,
  coalesce(p.shot_on, e.event_date)      as shot_on,
  p.photographer_actor_id,
  a.display_name                         as photographer,
  a.instagram                            as photographer_instagram,
  coalesce(p.credit_line, a.display_name) as credit,
  (p.photographer_actor_id is null)      as needs_credit
from public.event_photos p
left join public.events e on e.id = p.event_id and e.deleted_at is null
left join public.actors a on a.id = p.photographer_actor_id;

commit;

-- DOWN:
--   drop view public.v_photo_credits;
--   drop index public.idx_event_photos_photographer;
--   alter table public.event_photos
--     drop column photographer_actor_id, drop column credit_line, drop column shot_on;
