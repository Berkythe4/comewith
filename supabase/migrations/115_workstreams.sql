-- =============================================================================
-- 115_workstreams.sql  +  make meetings a shared, multi-user tool
--
-- Two changes:
--
-- 1. WORKSTREAMS — the second workflow. Tasks are discrete follow-ups with a
--    done state. A workstream is the other thing a meeting produces: an ongoing
--    thread ONE person drives and reports back on ("Martin ramping up
--    production", "the Green Room deal", "radio cadence"). It never completes;
--    it gets touch-base UPDATES until parked or absorbed.
--
-- 2. VISIBILITY — Keith, Martin and Henry each use this for their own syncs, and
--    for each other, so a meeting/workstream is PRIVATE to its creator by
--    default and gets "pushed to team" explicitly. Spinning a note into a task
--    puts THAT item on the shared Tasks board — the per-item way to push one
--    thing for all to see without sharing the whole meeting.
--
-- INSERT..RETURNING trap (097): the SELECT policy carries `created_by =
-- auth.uid()` DIRECTLY so a creator can read back their own just-inserted row.
-- =============================================================================
begin;

-- ---- visibility on meetings -------------------------------------------------
alter table public.meetings add column if not exists visibility text not null default 'private'
  check (visibility in ('private', 'team'));

-- Row-local predicate in the SELECT/USING clause (not a re-querying helper), so
-- .insert().select() works for the creator and teammates see only 'team' rows.
drop policy if exists "Admins manage meetings" on public.meetings;
create policy "Meetings owner or team read" on public.meetings for select
  using (public.is_admin() and (created_by = auth.uid() or visibility = 'team'));
create policy "Meetings owner write" on public.meetings for insert
  with check (public.is_admin() and created_by = auth.uid());
create policy "Meetings owner update" on public.meetings for update
  using (public.is_admin() and created_by = auth.uid())
  with check (public.is_admin() and created_by = auth.uid());
create policy "Meetings owner delete" on public.meetings for delete
  using (public.is_admin() and created_by = auth.uid());

-- Notes follow their meeting: you see a meeting's notes if you can see the
-- meeting; you can write/edit notes on a meeting you own.
drop policy if exists "Admins manage meeting_notes" on public.meeting_notes;
create policy "Meeting notes read" on public.meeting_notes for select
  using (public.is_admin() and exists (
    select 1 from public.meetings m where m.id = meeting_id
      and (m.created_by = auth.uid() or m.visibility = 'team')));
create policy "Meeting notes write" on public.meeting_notes for all
  using (public.is_admin() and exists (
    select 1 from public.meetings m where m.id = meeting_id and m.created_by = auth.uid()))
  with check (public.is_admin() and exists (
    select 1 from public.meetings m where m.id = meeting_id and m.created_by = auth.uid()));

-- ---- workstreams ------------------------------------------------------------
create table if not exists public.workstreams (
  id             uuid primary key default gen_random_uuid(),
  title          text not null,
  owner_actor_id uuid references public.actors(id) on delete set null,
  owner_label    text,
  status         text not null default 'active' check (status in ('active', 'blocked', 'parked', 'done')),
  visibility     text not null default 'private' check (visibility in ('private', 'team')),
  -- The strategic pillar this thread ladders up into — the SAME buckets the
  -- Strategy board and KPIs use, so daily work → workstream → pillar → vision is
  -- one chain, not a parallel taxonomy. 'radio'/'ops' extend the four for
  -- threads that don't map cleanly. Free text (no CHECK) so it stays flexible.
  pillar         text,
  next_touch     date,
  summary        text,
  event_id       uuid references public.events(id) on delete set null,
  created_by     uuid references auth.users(id) on delete set null,
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now(),
  deleted_at     timestamptz
);
create index if not exists workstreams_owner_idx on public.workstreams (owner_actor_id) where deleted_at is null;

create table if not exists public.workstream_updates (
  id            uuid primary key default gen_random_uuid(),
  workstream_id uuid not null references public.workstreams(id) on delete cascade,
  body          text not null,
  meeting_id    uuid references public.meetings(id) on delete set null,
  created_by    uuid references auth.users(id) on delete set null,
  created_at    timestamptz not null default now()
);
create index if not exists workstream_updates_ws_idx on public.workstream_updates (workstream_id, created_at desc);

-- Cross-links: a task can belong to a workstream; a meeting note can spawn one.
alter table public.tasks add column if not exists workstream_id uuid references public.workstreams(id) on delete set null;
alter table public.meeting_notes add column if not exists workstream_id uuid references public.workstreams(id) on delete set null;

alter table public.workstreams enable row level security;
alter table public.workstream_updates enable row level security;

create policy "Workstreams owner or team read" on public.workstreams for select
  using (public.is_admin() and (created_by = auth.uid() or visibility = 'team'));
create policy "Workstreams owner write" on public.workstreams for insert
  with check (public.is_admin() and created_by = auth.uid());
create policy "Workstreams owner update" on public.workstreams for update
  using (public.is_admin() and created_by = auth.uid())
  with check (public.is_admin() and created_by = auth.uid());
create policy "Workstreams owner delete" on public.workstreams for delete
  using (public.is_admin() and created_by = auth.uid());

create policy "Workstream updates read" on public.workstream_updates for select
  using (public.is_admin() and exists (
    select 1 from public.workstreams w where w.id = workstream_id
      and (w.created_by = auth.uid() or w.visibility = 'team')));
create policy "Workstream updates write" on public.workstream_updates for all
  using (public.is_admin() and exists (
    select 1 from public.workstreams w where w.id = workstream_id and w.created_by = auth.uid()))
  with check (public.is_admin() and exists (
    select 1 from public.workstreams w where w.id = workstream_id and w.created_by = auth.uid()));

revoke all on public.workstreams from anon;
revoke all on public.workstream_updates from anon;

drop trigger if exists audit_workstreams on public.workstreams;
create trigger audit_workstreams after insert or update or delete on public.workstreams
  for each row execute function public.audit_trigger_function();

commit;
