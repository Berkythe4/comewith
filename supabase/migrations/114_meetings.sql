-- =============================================================================
-- 114_meetings.sql
-- Meeting tracker: capture notes live during a meeting, and spin any note line
-- into a follow-up task with one click. The notes stay linked to the meeting,
-- and a task created from a note carries a back-reference so "where did this
-- come from?" is always answerable.
--   meetings        — one row per meeting (title, date, who, summary)
--   meeting_notes   — the running note lines; each can become a task
--   tasks.meeting_id / tasks.meeting_note_id — provenance of a follow-up
-- Team-visible working material: admin-only RLS, anon revoked. Not financial,
-- so no release gate — any team member with dashboard access sees meetings.
-- =============================================================================
begin;

create table if not exists public.meetings (
  id           uuid primary key default gen_random_uuid(),
  title        text not null,
  meeting_date date not null default current_date,
  attendees    text,                       -- free text; actor linkage can come later
  attendee_ids uuid[] default '{}',        -- optional structured attendees (actors)
  summary      text,
  event_id     uuid references public.events(id) on delete set null,
  status       text not null default 'open' check (status in ('open', 'closed')),
  created_by   uuid references auth.users(id) on delete set null,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now(),
  deleted_at   timestamptz
);

create table if not exists public.meeting_notes (
  id           uuid primary key default gen_random_uuid(),
  meeting_id   uuid not null references public.meetings(id) on delete cascade,
  body         text not null,
  -- What KIND of line this is, so a meeting reads back as structure not a blob.
  kind         text not null default 'note' check (kind in ('note', 'decision', 'action', 'question')),
  task_id      uuid references public.tasks(id) on delete set null,  -- set when spun off
  sort         int not null default 0,
  created_by   uuid references auth.users(id) on delete set null,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);
create index if not exists meeting_notes_meeting_idx on public.meeting_notes (meeting_id, sort);

-- A task can point back at the meeting / note it came out of.
alter table public.tasks add column if not exists meeting_id      uuid references public.meetings(id) on delete set null;
alter table public.tasks add column if not exists meeting_note_id uuid references public.meeting_notes(id) on delete set null;

-- 'meeting' is a new task source (a follow-up spun off a meeting note). The
-- existing CHECK would reject it, so widen it.
alter table public.tasks drop constraint if exists tasks_source_check;
alter table public.tasks add constraint tasks_source_check
  check (source = any (array['manual', 'template', 'jennifer_import', 'meeting']));

alter table public.meetings enable row level security;
alter table public.meeting_notes enable row level security;
drop policy if exists "Admins manage meetings" on public.meetings;
create policy "Admins manage meetings" on public.meetings for all
  using (public.is_admin()) with check (public.is_admin());
drop policy if exists "Admins manage meeting_notes" on public.meeting_notes;
create policy "Admins manage meeting_notes" on public.meeting_notes for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.meetings from anon;
revoke all on public.meeting_notes from anon;

-- Audit trigger, same pattern as 107 (reads the pk out of jsonb, so it's safe).
drop trigger if exists audit_meetings on public.meetings;
create trigger audit_meetings after insert or update or delete on public.meetings
  for each row execute function public.audit_trigger_function();
drop trigger if exists audit_meeting_notes on public.meeting_notes;
create trigger audit_meeting_notes after insert or update or delete on public.meeting_notes
  for each row execute function public.audit_trigger_function();

-- Nav module. Pinned next to Calendar & Tasks, master + operations/full.
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('meetings', 'Meetings', 'Pinned', 2, true, true, false, array['operations', 'marketing', 'full'])
on conflict (key) do update set built = true, signed_off = true, nav_group = 'Pinned';

commit;
