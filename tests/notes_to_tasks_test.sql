-- =============================================================================
-- notes_to_tasks_test.sql — gate for migration 139.
-- Runs 139 inside BEGIN..ROLLBACK: nothing persists. Proves the widened source
-- CHECK admits 'note' without rejecting anything it already allowed, that the
-- back-link survives deleting its note, and that 'site' needs no migration.
-- =============================================================================
begin;

-- ---- the migration body (139) ----------------------------------------------
alter table public.tasks drop constraint if exists tasks_source_check;
alter table public.tasks add constraint tasks_source_check
  check (source = any (array['manual', 'template', 'jennifer_import', 'meeting', 'note']));

alter table public.tasks
  add column if not exists feedback_note_id uuid references public.feedback_log(id) on delete set null;

create index if not exists tasks_feedback_note_idx
  on public.tasks (feedback_note_id) where feedback_note_id is not null;

-- ---- functional checks ------------------------------------------------------
create temp table t_res (label text, ok boolean, detail text) on commit drop;

do $do$
declare
  n_id uuid; t_id uuid; still uuid; bad boolean := false;
begin
  insert into public.feedback_log (type, note) values ('idea', '[TEST 139] convert me')
    returning id into n_id;

  -- 1. a task may now declare source 'note' and point back at it
  insert into public.tasks (title, status, source, pillar, feedback_note_id)
    values ('[TEST 139] converted', 'todo', 'note', 'site', n_id) returning id into t_id;
  insert into t_res values ('source=note accepted', t_id is not null, coalesce(t_id::text, 'NULL'));

  -- 2. 'site' is just text — 116 left pillar unconstrained on purpose
  insert into t_res values ('pillar=site stored', exists (
    select 1 from public.tasks where id = t_id and pillar = 'site'), 'site');

  -- 3. the pre-existing vocabulary still works (a widen must not narrow)
  begin
    insert into public.tasks (title, status, source) values ('[TEST 139] legacy', 'todo', 'meeting');
    insert into t_res values ('source=meeting still ok', true, 'meeting');
  exception when check_violation then
    insert into t_res values ('source=meeting still ok', false, 'REJECTED');
  end;

  -- 4. a bogus source is still refused
  begin
    insert into public.tasks (title, status, source) values ('[TEST 139] bogus', 'todo', 'nonsense');
    bad := true;
  exception when check_violation then
    bad := false;
  end;
  insert into t_res values ('bogus source rejected', not bad, case when bad then 'ACCEPTED' else 'rejected' end);

  -- 5. deleting the note must NOT delete the work it became
  delete from public.feedback_log where id = n_id;
  select id into still from public.tasks where id = t_id;
  insert into t_res values ('task survives note delete', still is not null, coalesce(still::text, 'GONE'));
  insert into t_res values ('back-link nulled, not cascaded', exists (
    select 1 from public.tasks where id = t_id and feedback_note_id is null), 'set null');
end
$do$;

select label, ok, detail from t_res
union all
select 'column added', count(*) = 1, count(*)::text
  from information_schema.columns
  where table_schema = 'public' and table_name = 'tasks' and column_name = 'feedback_note_id'
union all
select 'index created', count(*) = 1, count(*)::text
  from pg_indexes where tablename = 'tasks' and indexname = 'tasks_feedback_note_idx'
union all
select 'notes untouched', count(*) = 30, count(*)::text
  from public.feedback_log where note not like '[TEST 139]%';

rollback;
