-- =============================================================================
-- 080_sc_oauth.sql
-- SoundCloud OAuth (Authorization Code + PKCE) so the in-app station can be
-- pushed to a REAL playlist in the connected SoundCloud account (Artist Pro).
-- Single-row token store (id='singleton'); tokens are service-role material —
-- table is admin-RLS and the dashboard never selects tokens directly (it goes
-- through the sc-connect edge fn). Grants inherited from 013.
-- =============================================================================
begin;

create table if not exists public.sc_oauth (
  id text primary key default 'singleton',
  state text,                 -- CSRF + PKCE lookup during the auth handshake
  code_verifier text,         -- PKCE verifier, cleared after token exchange
  access_token text,
  refresh_token text,         -- single-use; rotated on every refresh
  expires_at timestamptz,
  sc_user_id text,
  sc_username text,
  connected_at timestamptz,
  updated_at timestamptz not null default now()
);
alter table public.sc_oauth enable row level security;
drop policy if exists "Admins read sc_oauth" on public.sc_oauth;
create policy "Admins read sc_oauth" on public.sc_oauth for all using (public.is_admin()) with check (public.is_admin());

-- Remember which SoundCloud playlist each in-app station maps to (create once, then update).
alter table public.sc_playlists add column if not exists sc_playlist_id text;
alter table public.sc_playlists add column if not exists sc_playlist_url text;

commit;
-- POST: token store ready (empty until Connect); sc_playlists can remember its
-- exported SoundCloud playlist. Secrets SC_CLIENT_ID / SC_CLIENT_SECRET set later.
