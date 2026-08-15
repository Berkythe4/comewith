-- =============================================================================
-- 139_notes_to_tasks.sql
-- "Convert to task" on the Notes page: a note that turns out to be real work
-- becomes a task and the note closes, so the same item stops living in two
-- places with two different states.
--
-- The 'site' bucket needs NO migration. 116 made `pillar` free text on purpose
-- ("no CHECK, to stay flexible; the UI offers the known set"), so adding a
-- bucket is a UI change — see WS_PILLARS in dashboard.html.
--
-- Additive. No policy change: tasks' existing RLS already governs who may
-- create one, and this adds no new access path.
-- =============================================================================
begin;

-- A task can now originate from a note. 114 made exactly this move to admit
-- 'meeting'; the CHECK has to be widened or the insert is rejected.
alter table public.tasks drop constraint if exists tasks_source_check;
alter table public.tasks add constraint tasks_source_check
  check (source = any (array['manual', 'template', 'jennifer_import', 'meeting', 'note']));

-- Back-link, mirroring tasks.meeting_note_id. Without it the note simply goes
-- 'done' and the trail from "we decided to do this" to "here is the work" is
-- gone. ON DELETE SET NULL so removing a note never deletes real work.
alter table public.tasks
  add column if not exists feedback_note_id uuid references public.feedback_log(id) on delete set null;

create index if not exists tasks_feedback_note_idx
  on public.tasks (feedback_note_id)
  where feedback_note_id is not null;

commit;
