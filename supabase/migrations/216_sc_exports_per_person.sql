-- 216: anyone can export a station to THEIR OWN SoundCloud, from the dashboard too.
--
-- 215 built the one-shot export for guest DJs (dj.html). The same flow now serves
-- the dashboard, so Keith, Martin, Henry each get the playlist in their own
-- account, and two people can export the same station independently. The stored
-- singleton connection (Keith's, which also powers "sync back") is untouched.
--   origin        where the export started, so sc-oauth sends them back there
--   requested_by  the logged-in admin (dashboard only; a guest DJ has no login)
-- Additive only: existing rows are all dj exports and take the default.

begin;

alter table public.sc_dj_exports
  add column origin text not null default 'dj' check (origin in ('dj', 'dashboard')),
  add column requested_by uuid references public.profiles(id) on delete set null;

create index sc_dj_exports_requested_by on public.sc_dj_exports(requested_by, completed_at desc)
  where requested_by is not null;

commit;
