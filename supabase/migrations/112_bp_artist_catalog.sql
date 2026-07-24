-- =============================================================================
-- 112_bp_artist_catalog.sql
-- Beatport's catalogue for an artist, cached per artist name.
-- Why: SoundCloud under-serves label artists — Claude VonStroke's profile
-- declares 147 tracks and hands an anonymous client 27, and every paging
-- strategy tested returns the same 27 (111). Beatport is where those releases
-- are. Keyed on a normalised artist NAME, not a SoundCloud URL, so an artist
-- with no SoundCloud can still have a catalogue.
-- Admin-only, anon-revoked. Beatport tokens are never stored here (they live in
-- beatport_oauth, only until their own JWT exp — see CLAUDE.md).
-- =============================================================================
begin;

create table if not exists public.bp_artist_catalog (
  artist_key    text primary key,
  artist_name   text not null,
  bp_artist_id  int,
  tracks        jsonb not null default '[]'::jsonb,
  track_count   int not null default 0,
  method        text,
  fetched_at    timestamptz not null default now()
);

alter table public.bp_artist_catalog enable row level security;
drop policy if exists "Admins manage bp_artist_catalog" on public.bp_artist_catalog;
create policy "Admins manage bp_artist_catalog" on public.bp_artist_catalog for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.bp_artist_catalog from anon;

commit;
