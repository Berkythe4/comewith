-- =============================================================================
-- 210_public_link_emphasis.sql
-- 209 added link_items.emphasis but did not add it to v_public_link_items, which
-- is the ONLY thing links.html can read. The flag was therefore set on prod and
-- invisible to the page — a column nobody could see is the same as no column.
--
-- Separate migration rather than an edit to 209: 209 is applied and its sha is
-- recorded, and rewriting an applied file is exactly the drift the apply
-- discipline exists to prevent.
--
-- create-or-replace can APPEND a column to a view (not drop or reorder one), so
-- emphasis goes last and the existing anon grant carries over untouched.
-- =============================================================================
begin;

create or replace view public.v_public_link_items as
  select i.id, p.slug as page_slug, i.label, i.url, i.subtitle, i.icon,
         i.thumb_url, i.style, i.sort_order, i.created_at, i.emphasis
    from public.link_items i
    join public.link_pages p on p.id = i.page_id
   where p.is_published = true
     and i.is_active    = true
     and (i.starts_at is null or i.starts_at <= now())
     and (i.ends_at   is null or i.ends_at   >= now());

notify pgrst, 'reload schema';

commit;
