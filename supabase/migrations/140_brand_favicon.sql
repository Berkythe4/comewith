-- 140: a brand.favicon key so the site can carry its own tab icon.
--
-- The dashboard has linked /icons/favicon-32.png and /icons/apple-touch-icon.png
-- since it shipped; the PUBLIC pages never linked anything, so comewith.org has
-- been rendering the browser's blank default in the tab. The icon files already
-- exist in /icons — they were only ever referenced by dashboard.html and sw.js.
--
-- Two halves to fixing that, and this is the DB half. The static <link> tags are
-- the front-end half and cover the default; this key is the OVERRIDE, so Keith
-- can swap the tab icon from the Site Editor without a code change — same shape
-- as brand.logo, which is already an uploaded URL in this table.
--
-- Seeded EMPTY on purpose. The Site Editor renders one field per row that exists
-- in site_content, so the row has to be here for the picker to appear at all;
-- an empty value means "no override" and the pages fall through to the static
-- /icons/favicon-32.png default.
--
-- Additive. One row, no schema/policy/grant change.

insert into public.site_content (key, value) values ('brand.favicon', '')
on conflict (key) do nothing;

-- POST: select key, value from public.site_content where key = 'brand.favicon';
--       -> exactly one row, value '' (or whatever has since been uploaded).
