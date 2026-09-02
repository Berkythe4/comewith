-- =============================================================================
-- 209_link_social_emphasis.sql
-- Social links on a links page are the point of the page, not a footnote.
--
-- They rendered as 42px icon-only circles under the bio — the easiest thing on
-- the page to scroll past — while YouTube and SoundCloud are the audience-growth
-- targets and Instagram is the main following page. Worse, YouTube and
-- SoundCloud were not on the page AT ALL: 207 seeded only Instagram and email.
--
-- WHY A COLUMN AND NOT A HARDCODED PLATFORM LIST. The renderer could simply
-- decide that youtube/soundcloud/instagram are always big. That is right for
-- Come With today and wrong the moment /links/keith wants Bandcamp first, or a
-- Dance Infusion page wants the donate link loudest. Emphasis is a property of
-- THIS link on THIS page, so it is a column, and the editor gets a toggle.
--
-- The URLs seeded below are all verified, not guessed:
--   youtube.com/@comewithnyc      — already linked from watch.html
--   soundcloud.com/comewithnyc    — the account six published mixes live under
--   instagram.com/comewithnyc     — site_content['contact.ig_url']
--
-- Additive: one nullable-with-default column plus rows. Nothing is dropped, so
-- this may ship ahead of its UI.
-- =============================================================================
begin;

alter table public.link_items
  add column if not exists emphasis text not null default 'normal'
  check (emphasis in ('normal', 'primary'));

comment on column public.link_items.emphasis is
  'primary = render this link big. On social rows that means a branded tile with the platform name and its action verb, instead of a small icon circle. Per link, because which platform matters is a per-page decision.';

-- Promote the three that carry the goal: YouTube and SoundCloud for growth,
-- Instagram for following. Keith had already added all three by hand through the
-- editor before this shipped, so on prod this is the whole change — it only sets
-- the flag, and touches neither his labels nor his ordering.
update public.link_items i
   set emphasis = 'primary'
  from public.link_pages p
 where p.id = i.page_id and p.slug = 'main'
   and i.style = 'social'
   and i.icon in ('youtube', 'soundcloud', 'instagram')
   and i.emphasis <> 'primary';

-- Only if a page is missing them entirely (a fresh environment, or a new page).
-- A no-op on prod: the guard below finds Keith's rows. URLs verified, not guessed
-- — see the header.
insert into public.link_items (page_id, label, url, subtitle, icon, style, emphasis, sort_order)
select p.id, v.label, v.url, v.subtitle, v.icon, 'social', 'primary', v.ord
  from public.link_pages p
  cross join lateral (values
      ('YouTube',    'https://www.youtube.com/@comewithnyc', 'Subscribe', 'youtube',    60),
      ('SoundCloud', 'https://soundcloud.com/comewithnyc',   'Follow',    'soundcloud', 70)
    ) as v(label, url, subtitle, icon, ord)
 where p.slug = 'main'
   and not exists (
     select 1 from public.link_items i
      where i.page_id = p.id and i.style = 'social' and i.icon = v.icon);

notify pgrst, 'reload schema';

commit;
