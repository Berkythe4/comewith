-- ============================================================
-- COME WITH — 184 a recap video can be staged before it is public
--
-- THE BUG THIS FIXES. The public pages filter recap videos with a pattern match
-- and nothing else:
--     a.filter(v => v && (ytId(v.url) || /soundcloud\.com|snd\.sc/i.test(v.url)))
-- Any string containing soundcloud.com passes. So a PRIVATE track was rendered
-- as a tile with a label and a dead player inside - not hidden, not an error,
-- just silence. A private YouTube video rendered "Video unavailable", and since
-- events.youtube_url also drives the homepage card thumbnail, a broken image on
-- the front page too.
--
-- Worse, the editor said otherwise. When resolve-media flagged a private link on
-- save, the confirm read "They'll be stored but stay hidden on the site until
-- fixed." Nothing enforced that. They were stored and they rendered. 184 makes
-- that sentence true.
--
-- THE SHAPE. Each entry in events.recap_videos gets an optional is_public flag,
-- and v_public_recap keeps only the entries where it is not false.
--
-- A MISSING FLAG MEANS PUBLIC. That is deliberate and it is the whole reason no
-- backfill is needed: every video on the site today keeps rendering exactly as
-- it does now. 175 flipped event_photos.is_public to default false and
-- explicitly left existing rows alone for the same reason - silently
-- un-publishing live content is a worse surprise than the old default was. New
-- entries are written with an explicit flag by the editor, and anything
-- resolve-media cannot embed is written staged automatically.
--
-- WHY HERE AND NOT content_assets. content_assets (121) has an is_public column
-- and looks like the right home, but nothing public reads that table - the flag
-- is decorative and all 9 rows have it false. The public contract is
-- v_public_recap over events.recap_videos, and that is the one place all three
-- public pages (index / watch / event) go through. Fixing it here fixes all
-- three at once.
-- ============================================================
begin;

create or replace view public.v_public_recap as
select e.id,
       e.name,
       e.event_date,
       v.name as venue_name,
       e.series,
       e.type,
       e.hero_image_path,
       -- The card thumbnail comes from the first PUBLIC YouTube link. A staged
       -- video must not leave a broken image on the homepage. When there are no
       -- recap_videos at all, fall back to the legacy column untouched.
       case
         when coalesce(jsonb_array_length(e.recap_videos), 0) = 0 then e.youtube_url
         else (
           select t.v ->> 'url'
             from jsonb_array_elements(e.recap_videos) with ordinality as t(v, ord)
            where coalesce((t.v ->> 'is_public')::boolean, true)
              and t.v ->> 'url' ~* '(youtube\.com|youtu\.be)'
            order by t.ord
            limit 1)
       end as youtube_url,
       e.recap_blurb,
       -- Staged entries are removed server-side rather than filtered in three
       -- separate pages, so there is one place this can be got wrong.
       coalesce((
         select jsonb_agg(t.v order by t.ord)
           from jsonb_array_elements(coalesce(e.recap_videos, '[]'::jsonb)) with ordinality as t(v, ord)
          where coalesce((t.v ->> 'is_public')::boolean, true)
       ), '[]'::jsonb) as recap_videos
  from public.events e
  left join public.venues v on v.id = e.venue_id
 where e.is_featured = true and e.deleted_at is null
 order by e.event_date desc;

-- v_public_recap is anon-readable BY DESIGN (it is the public recap feed); the
-- create-or-replace above preserves that grant. Re-asserted so a future reader
-- does not "fix" it: this one is supposed to be readable.
grant select on public.v_public_recap to anon;

notify pgrst, 'reload schema';

commit;

-- DOWN: restore 063's definition of v_public_recap.
