-- =============================================================================
-- 221_content_multi_artist.sql — tag MANY artists on every kind of event content
-- =============================================================================
-- The event hub's Content tab holds four kinds of thing, and until now only two
-- could name an artist, and only one each:
--
--   recap videos     events.recap_videos[].artist_id   (jsonb, one id)
--   library assets   content_assets.artist_id          (one id)
--   social posts     -                                  (none)
--   photos           -                                  (none; photographer is a CREDIT, not a tag)
--
-- A b2b set, a crowd shot with three DJs in it, a recap reel cut from a whole
-- lineup - all of it is about more than one person. This adds `artist_ids` to
-- every kind:
--
--   * content_assets / social_posts / event_photos get `artist_ids uuid[]`
--     (not null, default '{}', no null elements). Arrays, not a join table,
--     because recap videos are jsonb items with no id of their own to join to;
--     one shape everywhere keeps the dashboard to one writer. Actors are
--     soft-deleted, so a dangling id cannot appear by deletion.
--   * recap items get an `artist_ids` json array. The old single `artist_id`
--     key is KEPT and the dashboard keeps it equal to artist_ids[0], so nothing
--     that still reads the singular key goes blind. Nothing is rewritten here:
--     the view reads both keys, and an item is upgraded the next time it saves.
--   * content_assets.artist_id likewise stays, mirrored to artist_ids[0].
--
-- ADDITIVE. The deployed dashboard ignores the new columns and keeps working.
--
-- v_artist_content (the public artist page's "Content" list) now expands
-- artist_ids, and ALSO stops listing a STAGED recap item. It never checked
-- is_public, so tagging a still-private track would have published it on the
-- artist's profile while v_public_recap kept it off the event page. 0 rows are
-- in that state today (checked 2026-10-17), so the gate changes nothing live.
-- Photos/posts/assets stay internal: tagging them does not publish anything.

begin;

alter table public.content_assets
  add column if not exists artist_ids uuid[] not null default '{}';
alter table public.social_posts
  add column if not exists artist_ids uuid[] not null default '{}';
alter table public.event_photos
  add column if not exists artist_ids uuid[] not null default '{}';

alter table public.content_assets
  add constraint content_assets_artist_ids_no_nulls check (array_position(artist_ids, null) is null);
alter table public.social_posts
  add constraint social_posts_artist_ids_no_nulls check (array_position(artist_ids, null) is null);
alter table public.event_photos
  add constraint event_photos_artist_ids_no_nulls check (array_position(artist_ids, null) is null);

-- "everything this artist is in" is the query these exist for
create index if not exists content_assets_artist_ids_gin on public.content_assets using gin (artist_ids);
create index if not exists social_posts_artist_ids_gin   on public.social_posts   using gin (artist_ids);
create index if not exists event_photos_artist_ids_gin   on public.event_photos   using gin (artist_ids);

-- carry the existing single tag across (0 rows on prod today; harmless if more)
update public.content_assets
   set artist_ids = array[artist_id]
 where artist_id is not null and artist_ids = '{}';

-- Same columns, same order, so CREATE OR REPLACE keeps the existing grants.
create or replace view public.v_artist_content as
  select x.actor_id,
         t.v ->> 'url'   as url,
         t.v ->> 'label' as label,
         e.id as event_id, e.name as event_name, e.event_date
  from public.events e
  cross join lateral jsonb_array_elements(coalesce(e.recap_videos, '[]'::jsonb)) with ordinality t(v, ord)
  cross join lateral (
    select a::uuid as actor_id
      from jsonb_array_elements_text(
             case when jsonb_typeof(t.v -> 'artist_ids') = 'array' then t.v -> 'artist_ids' else '[]'::jsonb end) a
     where a <> ''
    union                                   -- dedupes the mirrored legacy key
    select (t.v ->> 'artist_id')::uuid
     where coalesce(t.v ->> 'artist_id', '') <> ''
  ) x
  where e.deleted_at is null
    -- missing flag = public, exactly as v_public_recap reads it
    and coalesce((t.v ->> 'is_public')::boolean, true);

notify pgrst, 'reload schema';

commit;

-- DOWN:
--   restore 065's v_artist_content;
--   alter table public.content_assets drop column artist_ids;
--   alter table public.social_posts  drop column artist_ids;
--   alter table public.event_photos  drop column artist_ids;
