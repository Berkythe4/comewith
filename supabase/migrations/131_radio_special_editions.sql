-- 131: special-edition episode series (e.g. a festival run of daily drops instead
-- of the weekly). Episodes in an edition share edition_name; edition_seq is the
-- day/order within it. Normal weekly episodes leave both null.
alter table public.sc_playlists
  add column if not exists edition_name text,
  add column if not exists edition_seq  integer;
create index if not exists idx_sc_playlists_edition on public.sc_playlists (edition_name) where edition_name is not null;
