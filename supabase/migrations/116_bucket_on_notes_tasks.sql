-- =============================================================================
-- 116_bucket_on_notes_tasks.sql
-- A "bucket" (strategic pillar) on meeting notes and tasks, so a line is
-- labelled ONCE and the label carries through when it becomes a task or a
-- workstream. Same taxonomy as workstreams.pillar and the Strategy board's
-- workstreams (content/audience/parties/dance_infusion + radio/ops), so every
-- piece of work — note, task, workstream — rolls up into the same buckets.
-- Free text (no CHECK) to stay flexible; the UI offers the known set.
-- =============================================================================
begin;
alter table public.meeting_notes add column if not exists pillar text;
alter table public.tasks         add column if not exists pillar text;
create index if not exists tasks_pillar_idx on public.tasks (pillar) where pillar is not null;
commit;
