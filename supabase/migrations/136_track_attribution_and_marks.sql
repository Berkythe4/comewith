-- 136_track_attribution_and_marks.sql
-- Multiple people now build a station from the same pool of songs, so a row needs
-- to say WHO put it there and WHERE each person stands on it.
--
-- Two parts, both scoped to the episode (decision 2026-07-30): a song's role in a
-- set is per-set, so nothing here touches sc_song_log — the permanent song memory
-- stays about played/passed/carried only.
--
--  1. sc_playlist_tracks.added_by  — automatic attribution, defaults to auth.uid()
--  2. sc_track_marks               — one row per (track, person): claimed/maybe/veto
--
-- Why marks are a CHILD TABLE and not a column: the point is to surface
-- disagreement, which means Keith and Martin must be able to hold different
-- positions on the SAME track at the same time. A single column would let the
-- last writer silently overwrite the other person's call.
--
-- PRE : sc_playlist_tracks has no attribution; no marks table exists
-- POST: added_by column (nullable, set-null on profile delete) + sc_track_marks
--       with RLS: any admin READS all marks, but may only write their OWN.

-- (1) attribution ------------------------------------------------------------
alter table public.sc_playlist_tracks
  add column if not exists added_by uuid references public.profiles(id) on delete set null;

comment on column public.sc_playlist_tracks.added_by is
  'Who added this song to the station. Defaults to auth.uid(); NULL for rows added by a
   service-role caller (sc-connect sync/carry-over, dj-station) or predating migration 136.';

-- Default rather than a trigger: a trigger would have to invent a user for the
-- service-role callers, and "nobody in particular" is the honest answer there.
alter table public.sc_playlist_tracks
  alter column added_by set default auth.uid();

create index if not exists sc_playlist_tracks_added_by_idx
  on public.sc_playlist_tracks (added_by);

-- (2) per-person marks -------------------------------------------------------
create table if not exists public.sc_track_marks (
  id         uuid primary key default gen_random_uuid(),
  track_id   uuid not null references public.sc_playlist_tracks(id) on delete cascade,
  user_id    uuid not null references public.profiles(id) on delete cascade,
  mark       text not null check (mark in ('claimed', 'maybe', 'veto')),
  note       text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (track_id, user_id)          -- one position per person per track
);

create index if not exists sc_track_marks_track_idx on public.sc_track_marks (track_id);

alter table public.sc_track_marks enable row level security;

-- Read: any admin sees everyone's marks — seeing the disagreement is the feature.
-- Write: you may only create/change/remove YOUR OWN mark.
-- Note the 097 lesson does not bite here: the SELECT predicate is is_admin(), not
-- row-local, so .insert().select() can see its own new row mid-statement.
drop policy if exists sc_track_marks_select on public.sc_track_marks;
create policy sc_track_marks_select on public.sc_track_marks
  for select using (public.is_admin());

drop policy if exists sc_track_marks_write on public.sc_track_marks;
create policy sc_track_marks_write on public.sc_track_marks
  for insert with check (public.is_admin() and user_id = auth.uid());

drop policy if exists sc_track_marks_update on public.sc_track_marks;
create policy sc_track_marks_update on public.sc_track_marks
  for update using (public.is_admin() and user_id = auth.uid())
          with check (public.is_admin() and user_id = auth.uid());

drop policy if exists sc_track_marks_delete on public.sc_track_marks;
create policy sc_track_marks_delete on public.sc_track_marks
  for delete using (public.is_admin() and user_id = auth.uid());

-- Grants come from 013's ALTER DEFAULT PRIVILEGES — do NOT add a blanket anon grant
-- here (that is the 016/017 regression 019 had to fix). Belt-and-braces revoke:
revoke all on public.sc_track_marks from anon;

comment on table public.sc_track_marks is
  'Per-person position on a station track: claimed / maybe / veto. Admins read all,
   write only their own row. Episode-scoped (cascades with the track). Migration 136.';
