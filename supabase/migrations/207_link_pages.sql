-- =============================================================================
-- 207_link_pages.sql
-- Link-in-bio pages ("linktree"): comewith.org/links, /links/di, /links/keith.
--
-- WHY A TABLE AND NOT site_content KEYS
--   site_content is a flat key/value CMS - right for "the hero headline", wrong
--   for an ORDERED, GROWING list Keith reorders from his phone between sets.
--   Links need sort_order, a schedule window and an on/off switch, so they are
--   rows. A page's own chrome (title, bio, avatar, colours) is a handful of
--   fields per page, so it sits on link_pages rather than as a hundred
--   site_content keys nobody could keep straight.
--
-- WHY MULTIPLE PAGES
--   Come With, Dance Infusion and Keith-as-a-DJ point different audiences at
--   different links. One table addressed by slug means a fourth page is a row
--   Keith adds himself - not a migration, and not a conversation with Claude.
--
-- SECURITY SHAPE (the 030 pattern, deliberately)
--   The TABLES are anon-revoked and admin-RLS'd; the public site reads two
--   dedicated VIEWS exposing only display fields, for PUBLISHED pages and
--   ACTIVE, in-window links. 013's ALTER DEFAULT PRIVILEGES grants new tables to
--   anon automatically, so the revokes below are load-bearing, not decoration.
--   Neither view is security_invoker, so each runs as owner and its WHERE clause
--   is the whole gate - same as v_public_events (030) and v_artist_gigs (205).
--
-- PUBLISHING IS A DELIBERATE TOGGLE. is_published defaults false, so a page
-- being built is invisible until Keith says otherwise - the same posture as
-- photos.is_public and actors.public_profile.
--
-- Additive only: new tables, new views, one module_registry row. Nothing
-- existing is dropped or tightened, so this may ship ahead of its UI.
-- =============================================================================
begin;

-- -- Pages ---------------------------------------------------------------------
create table if not exists public.link_pages (
  id              uuid primary key default gen_random_uuid(),
  slug            text not null unique,
  title           text not null,
  tagline         text,
  avatar_url      text,
  footer_note     text,
  seo_description text,
  og_image_url    text,
  -- Theme is jsonb on purpose: the editor writes a known set of keys (bg, bg2,
  -- accent, text, dim, btn_style, btn_radius, font, align, avatar_shape,
  -- bg_image, bg_dim, preset) and the page falls back to the Come With palette
  -- for anything absent. A new knob is then a UI change, never a migration -
  -- which is the whole point of "customisable without Claude".
  theme           jsonb   not null default '{}'::jsonb,
  is_published    boolean not null default false,
  sort_order      integer not null default 0,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  updated_by      uuid references public.profiles(id),
  constraint link_pages_slug_shape check (slug ~ '^[a-z0-9][a-z0-9-]{0,39}$')
);

comment on table public.link_pages is
  'One link-in-bio page per row, addressed by slug (comewith.org/links/<slug>). Anon-revoked; the public site reads v_public_link_pages.';
comment on column public.link_pages.is_published is
  'False = invisible to the public, including through the views. Publishing is a deliberate toggle, never a side effect of editing.';

-- -- Links on a page ------------------------------------------------------------
create table if not exists public.link_items (
  id          uuid primary key default gen_random_uuid(),
  page_id     uuid not null references public.link_pages(id) on delete cascade,
  label       text not null,
  url         text,
  subtitle    text,
  icon        text,          -- an emoji, shown before the label
  thumb_url   text,          -- 'feature' rows only
  -- button  = the ordinary full-width row
  -- feature = a big card with a thumbnail (the thing you actually want clicked)
  -- header  = a section heading, no link at all
  -- social  = a small icon pill in the row under the bio
  style       text not null default 'button'
              check (style in ('button', 'feature', 'header', 'social')),
  is_active   boolean not null default true,
  -- A schedule window, so "Tickets - Sat" can be set once and disappear by
  -- itself. This is why the public list is a view: the window has to be applied
  -- server-side, or a link that should be gone still ships to the browser and is
  -- merely hidden by CSS.
  starts_at   timestamptz,
  ends_at     timestamptz,
  sort_order  integer not null default 0,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now(),
  -- A header has nothing to click; everything else must go somewhere.
  constraint link_items_url_required check (style = 'header' or coalesce(url, '') <> '')
);

create index if not exists link_items_page_idx on public.link_items (page_id, sort_order, created_at);

comment on table public.link_items is
  'Ordered links on a link_pages row. Anon-revoked; the public site reads v_public_link_items, which applies is_active and the starts_at/ends_at window.';

drop trigger if exists set_updated_at on public.link_pages;
create trigger set_updated_at before update on public.link_pages
  for each row execute function public.handle_updated_at();
drop trigger if exists set_updated_at on public.link_items;
create trigger set_updated_at before update on public.link_items
  for each row execute function public.handle_updated_at();

-- -- RLS: admin-only on the tables themselves ------------------------------------
alter table public.link_pages enable row level security;
alter table public.link_items enable row level security;
drop policy if exists "Admins manage link pages" on public.link_pages;
create policy "Admins manage link pages" on public.link_pages for all
  using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage link items" on public.link_items;
create policy "Admins manage link items" on public.link_items for all
  using (public.is_admin()) with check (public.is_admin());

revoke all on public.link_pages from anon;
revoke all on public.link_items from anon;

-- -- The two public views --------------------------------------------------------
create or replace view public.v_public_link_pages as
  select p.slug, p.title, p.tagline, p.avatar_url, p.footer_note,
         p.seo_description, p.og_image_url, p.theme
    from public.link_pages p
   where p.is_published = true;

create or replace view public.v_public_link_items as
  select i.id, p.slug as page_slug, i.label, i.url, i.subtitle, i.icon,
         i.thumb_url, i.style, i.sort_order, i.created_at
    from public.link_items i
    join public.link_pages p on p.id = i.page_id
   where p.is_published = true
     and i.is_active    = true
     and (i.starts_at is null or i.starts_at <= now())
     and (i.ends_at   is null or i.ends_at   >= now());

comment on view public.v_public_link_pages is
  'Anon-readable chrome for a PUBLISHED link-in-bio page. No internal fields, no unpublished rows.';
comment on view public.v_public_link_items is
  'Anon-readable links for PUBLISHED pages: active only, and only inside their starts_at/ends_at window.';

grant select on public.v_public_link_pages to anon, authenticated;
grant select on public.v_public_link_items to anon, authenticated;

-- -- Click stats (admin only) -----------------------------------------------------
-- links.html stamps every link with data-track="link:<uuid>", so the beacon's
-- link_label carries an EXACT identity. Matching on link_url instead would be a
-- guess: the beacon stores the resolved absolute URL while a link row may hold a
-- relative path, and a link whose URL was edited would silently inherit the old
-- one's history. site_events is admin-RLS'd but this view runs as owner, so the
-- revoke below is what keeps it admin-only.
create or replace view public.v_link_click_stats as
  select i.id                        as link_id,
         p.slug                      as page_slug,
         i.label,
         count(se.id) filter (where se.occurred_at >= now() - interval '30 days') as clicks_30d,
         count(se.id)                                                             as clicks_all,
         max(se.occurred_at)                                                      as last_click
    from public.link_items i
    join public.link_pages p on p.id = i.page_id
    left join public.site_events se
           on se.kind = 'click'
          and se.link_label = 'link:' || i.id::text
   group by i.id, p.slug, i.label;

revoke all on public.v_link_click_stats from anon;
grant select on public.v_link_click_stats to authenticated;

-- -- Nav module, next to Site Editor (the other public-site surface) --------------
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('links', 'Links Page', 'Team HQ', 55, true, true, false, array['marketing','full'])
on conflict (key) do update set built = true, label = excluded.label,
                                nav_group = excluded.nav_group, sort_order = excluded.sort_order;

-- -- Seed: the Come With page, UNPUBLISHED, with links that already exist --------
-- Seeded from what the site already says rather than invented, so nothing here
-- is a guess. Internal links stay relative - the page is served from the same
-- host - and they still count, because links.html tags every link for the beacon.
insert into public.link_pages (slug, title, tagline, theme, is_published)
select 'main', 'Come With', 'Parties, radio, and a room that feels like home. NYC.',
       jsonb_build_object('preset', 'comewith'), false
 where not exists (select 1 from public.link_pages where slug = 'main');

insert into public.link_items (page_id, label, url, subtitle, icon, style, sort_order)
select p.id, v.label, v.url, v.subtitle, v.icon, v.style, v.ord
  from public.link_pages p
  cross join lateral (values
      ('Come With Radio', '/radio.html', 'Every show, every tracklist', 'radio',  'feature', 10),
      ('Upcoming events', '/#events',    'Where we are next',           'ticket', 'button',  20),
      ('Watch',           '/watch.html', 'Recaps from recent rooms',    'play',   'button',  30),
      ('Book us',         '/#book',      'Parties, production, DJs',    'mail',   'button',  40)
    ) as v(label, url, subtitle, icon, style, ord)
 where p.slug = 'main'
   and not exists (select 1 from public.link_items where page_id = p.id);

-- Instagram / email come from the site's own contact fields, when they are set.
insert into public.link_items (page_id, label, url, icon, style, sort_order)
select p.id, 'Instagram', c.value, 'instagram', 'social', 100
  from public.link_pages p
  join public.site_content c on c.key = 'contact.ig_url'
 where p.slug = 'main' and coalesce(c.value, '') <> ''
   and not exists (select 1 from public.link_items i
                    where i.page_id = p.id and i.style = 'social' and i.label = 'Instagram');

insert into public.link_items (page_id, label, url, icon, style, sort_order)
select p.id, 'Email', 'mailto:' || c.value, 'mail', 'social', 110
  from public.link_pages p
  join public.site_content c on c.key = 'contact.email'
 where p.slug = 'main' and coalesce(c.value, '') <> ''
   and not exists (select 1 from public.link_items i
                    where i.page_id = p.id and i.style = 'social' and i.label = 'Email');

notify pgrst, 'reload schema';

commit;
