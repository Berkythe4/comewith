-- =============================================================================
-- 087_watchlist.sql
-- Watchlist: tag local artists to keep an eye on + shows worth attending, with a
-- reason code, so there's a growing archive of significant upcoming events to use
-- later. kind='artist' (ref = lowercased name) or 'show' (ref = ra_event id).
-- Admin RLS. Grants inherited from 013.
-- =============================================================================
begin;
create table if not exists public.watchlist (
  id uuid primary key default gen_random_uuid(),
  kind text not null check (kind in ('artist','show')),
  ref text not null,                 -- artist name (lower) or ra_events.ra_id
  label text not null,               -- display name
  sublabel text,                     -- socials (artist) / lineup (show)
  event_date date,                   -- for shows
  venue text,
  reason text,                       -- reason code (see UI)
  note text,
  archived boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (kind, ref)
);
create index if not exists idx_watchlist_kind on public.watchlist(kind, archived);
alter table public.watchlist enable row level security;
drop policy if exists "Admins manage watchlist" on public.watchlist;
create policy "Admins manage watchlist" on public.watchlist for all using (public.is_admin()) with check (public.is_admin());
commit;
-- POST: watchlist table (anon-blocked). Tag from Best Nights + Artist Radio.
