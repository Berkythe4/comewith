-- 215: guest DJs export an episode to THEIR OWN SoundCloud.
--
-- The dashboard's export writes to one stored connection (sc_oauth 'singleton',
-- Keith's account). A guest DJ working from dj.html?ep=<token> has no login, so
-- each export is a one-shot OAuth round trip: dj-station starts it (PKCE verifier
-- + state stored here), SoundCloud redirects to sc-oauth, which exchanges the code,
-- builds the playlist in the DJ's account through sc-connect, and records the
-- outcome here. The DJ's access token is used for that one request and NEVER
-- stored - same principle as Beatport (no standing credential at rest).
--
-- Service role writes it; admins may read it (who exported what, where it went).
-- Anon has no business here at all.

begin;

create table public.sc_dj_exports (
  state          text primary key,
  playlist_id    uuid not null references public.sc_playlists(id) on delete cascade,
  code_verifier  text,                       -- cleared once used
  created_at     timestamptz not null default now(),
  completed_at   timestamptz,
  ok             boolean,
  sc_username    text,                       -- whose account it landed in
  result_url     text,                       -- the new playlist
  tracks         int,
  skipped        jsonb,                      -- titles SoundCloud would not take
  error          text
);
create index sc_dj_exports_playlist on public.sc_dj_exports(playlist_id, created_at desc);

alter table public.sc_dj_exports enable row level security;
create policy sc_dj_exports_admin on public.sc_dj_exports for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.sc_dj_exports from anon;

commit;
