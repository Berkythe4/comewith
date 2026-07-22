-- =============================================================================
-- 103_beatport_and_anon_revoke.sql
--
-- (1) beatport_oauth — singleton token store for the in-app "check availability"
--     button. Mirrors sc_oauth (080) exactly: admin-RLS'd, anon revoked, tokens
--     read/written by the edge function through the service role. Beatport's
--     refresh tokens ROTATE on every use, so they must be persisted somewhere
--     writable at runtime — an env secret can't be rewritten by the function.
--     NOT site_content: that table is anon-readable (standing rule).
--
-- (2) Revoke anon on the two radio playlist tables. Both were created in 079 and
--     picked up table-level grants from the 013 default privileges. RLS has
--     always blocked the rows (an anon REST GET returns 200 [] — verified, no
--     data ever leaked), but the grant shouldn't be there: it's the same shape as
--     the 016/017 regression, and it means one over-broad policy away from a real
--     leak. sc_song_log + the listener_* tables were revoked properly in 099;
--     this brings 079's tables in line. After this, anon gets 401 not 200 [].
-- =============================================================================
begin;

-- (1) Beatport token store -----------------------------------------------------
create table if not exists public.beatport_oauth (
  id text primary key default 'singleton',
  access_token text,
  refresh_token text,          -- rotates on every refresh; rewritten in place
  expires_at timestamptz,
  last_error text,             -- surfaced in the dashboard when the chain breaks
  updated_at timestamptz not null default now()
);
alter table public.beatport_oauth enable row level security;
drop policy if exists "Admins manage beatport_oauth" on public.beatport_oauth;
create policy "Admins manage beatport_oauth" on public.beatport_oauth for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.beatport_oauth from anon;

-- Cache of what each station track was matched to, so the availability check
-- doesn't re-hit the APIs every time you open the panel.
alter table public.sc_playlist_tracks add column if not exists beatport_url text;
alter table public.sc_playlist_tracks add column if not exists beatport_price text;
alter table public.sc_playlist_tracks add column if not exists bandcamp_url text;
alter table public.sc_playlist_tracks add column if not exists sources_checked_at timestamptz;

-- (2) Close the anon grants on 079's tables ------------------------------------
revoke all on public.sc_playlists from anon;
revoke all on public.sc_playlist_tracks from anon;

commit;
-- POST: beatport_oauth (admin-only, anon-revoked); four availability-cache
-- columns on sc_playlist_tracks; anon can no longer even reach sc_playlists /
-- sc_playlist_tracks at the grant level (was 200 [] via RLS, now 401).
-- Public reads of stations stay function-only via get-station (service role).
