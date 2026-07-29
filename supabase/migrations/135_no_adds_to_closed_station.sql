-- 135_no_adds_to_closed_station.sql
-- Stop songs being added to a CLOSED episode. On 2026-07-29 "She The Last One"
-- was accidentally added to EP 1 (live, published) — it appeared on the public
-- episode page and the dashboard offered no way to take it off, because the
-- remove controls are gated to the working station.
--
-- Why a TRIGGER and not RLS: three dashboard paths insert as the signed-in user,
-- but dj-station and sc-connect insert with the SERVICE ROLE, which bypasses RLS
-- entirely. A trigger is the only guard that covers every path.
--
-- Closed = 'live' | 'archived'.  Still open: 'building', 'testing' (the working
-- station) and 'planned' (future skeletons the DJ link adds to by design).
-- Verified safe against both service-role callers before writing this:
--   * sc-connect sync-back inserts into the working station (building/testing)
--   * sc-connect carry-over inserts into a NEWLY created station (default 'building')
--
-- Removing/editing tracks on a live episode stays allowed on purpose — that's the
-- fix-a-mistake path. Only ADDING is closed off.
--
-- PRE : any caller can insert a track row against a live/archived playlist
-- POST: such an insert raises; message is user-legible because the dashboard
--       surfaces error.message directly in its toast.

create or replace function public.sc_tracks_block_closed()
returns trigger
language plpgsql
security definer            -- so the status lookup never depends on the caller's RLS
set search_path = public, pg_temp
as $$
declare
  st text;
  no int;
begin
  select status, station_no into st, no
    from public.sc_playlists where id = new.playlist_id;

  if st in ('live', 'archived') then
    raise exception
      'EP % is % — reopen the episode before adding songs (set its status back to testing).',
      coalesce(no::text, '?'), st
      using errcode = 'check_violation';
  end if;

  return new;
end;
$$;

drop trigger if exists trg_sc_tracks_block_closed on public.sc_playlist_tracks;
create trigger trg_sc_tracks_block_closed
  before insert or update of playlist_id on public.sc_playlist_tracks
  for each row execute function public.sc_tracks_block_closed();

comment on function public.sc_tracks_block_closed() is
  'Blocks INSERTs (and re-parenting UPDATEs) of tracks onto live/archived episodes. See migration 135.';
