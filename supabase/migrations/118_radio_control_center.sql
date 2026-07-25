-- =============================================================================
-- 118_radio_control_center.sql
-- Backing for the Artist Radio Control Center:
--   • tasks.station_id — link a production task to its EPISODE (sc_playlists),
--     so the center can show/scope tasks per episode. The weekly generator sets
--     it; nullable for everything else.
--   • station_notes — a per-episode team-communication thread (Keith ↔ Janelle
--     ↔ Martin ↔ Henry), like the workstream update log but tied to an episode.
-- Admin-only, anon revoked, audited. Not financial — no release gate.
-- =============================================================================
begin;

alter table public.tasks add column if not exists station_id uuid references public.sc_playlists(id) on delete set null;
create index if not exists tasks_station_idx on public.tasks (station_id) where station_id is not null;

create table if not exists public.station_notes (
  id          uuid primary key default gen_random_uuid(),
  station_id  uuid not null references public.sc_playlists(id) on delete cascade,
  body        text not null,
  created_by  uuid references auth.users(id) on delete set null,
  created_at  timestamptz not null default now()
);
create index if not exists station_notes_station_idx on public.station_notes (station_id, created_at desc);

alter table public.station_notes enable row level security;
drop policy if exists "Admins manage station_notes" on public.station_notes;
create policy "Admins manage station_notes" on public.station_notes for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.station_notes from anon;

drop trigger if exists audit_station_notes on public.station_notes;
create trigger audit_station_notes after insert or update or delete on public.station_notes
  for each row execute function public.audit_trigger_function();

commit;
