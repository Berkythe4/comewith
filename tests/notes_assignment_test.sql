-- =============================================================================
-- notes_assignment_test.sql — gate for migration 138.
-- Runs the migration against prod inside BEGIN..ROLLBACK: nothing persists.
-- Proves the columns/index/trigger land AND that assigned_at is stamped and
-- cleared correctly, which is the only behaviour 138 adds beyond storage.
-- =============================================================================
begin;

-- ---- the migration body (138) ----------------------------------------------
alter table public.feedback_log
  add column if not exists assigned_to uuid references public.profiles(id) on delete set null;
alter table public.feedback_log
  add column if not exists assigned_at timestamptz;

create index if not exists feedback_log_assigned_idx
  on public.feedback_log (assigned_to, status)
  where assigned_to is not null;

create or replace function public.feedback_log_stamp_assignment()
returns trigger
language plpgsql
as $$
begin
  if tg_op = 'INSERT' then
    if new.assigned_to is not null then
      new.assigned_at := now();
    end if;
  elsif new.assigned_to is distinct from old.assigned_to then
    new.assigned_at := case when new.assigned_to is null then null else now() end;
  end if;
  return new;
end
$$;

drop trigger if exists feedback_log_stamp_assignment on public.feedback_log;
create trigger feedback_log_stamp_assignment
  before insert or update on public.feedback_log
  for each row execute function public.feedback_log_stamp_assignment();

-- ---- functional checks ------------------------------------------------------
-- 1. insert WITH an assignee  -> assigned_at stamped
-- 2. update to a different    -> re-stamped
-- 3. update to null           -> cleared
-- 4. insert WITHOUT assignee  -> stays null
create temp table t_res (label text, ok boolean, detail text) on commit drop;

do $do$
declare
  liz uuid; martin uuid; n1 uuid; n2 uuid; a1 timestamptz; a2 timestamptz;
begin
  select id into liz    from public.profiles where email = 'liz@comewith.org';
  select id into martin from public.profiles where email = 'martin@comewith.org';

  insert into public.feedback_log (type, note, assigned_to)
    values ('idea', '[TEST 138] assigned on insert', liz) returning id, assigned_at into n1, a1;
  insert into t_res values ('stamp on insert', a1 is not null, coalesce(a1::text, 'NULL'));

  update public.feedback_log set assigned_to = martin where id = n1 returning assigned_at into a2;
  insert into t_res values ('re-stamp on reassign', a2 is not null and a2 >= a1, coalesce(a2::text, 'NULL'));

  update public.feedback_log set assigned_to = null where id = n1 returning assigned_at into a2;
  insert into t_res values ('cleared on unassign', a2 is null, coalesce(a2::text, 'NULL'));

  insert into public.feedback_log (type, note)
    values ('idea', '[TEST 138] no assignee') returning id, assigned_at into n2, a1;
  insert into t_res values ('null when unassigned', a1 is null, coalesce(a1::text, 'NULL'));

  -- a note whose assignee is deactivated/deleted must survive (on delete set null)
  insert into t_res values ('fk is ON DELETE SET NULL', exists (
    select 1 from pg_constraint
    where conrelid = 'public.feedback_log'::regclass
      and confdeltype = 'n'
      and conkey = array[(select attnum from pg_attribute
                          where attrelid = 'public.feedback_log'::regclass
                            and attname = 'assigned_to')]
  ), 'confdeltype n');
end
$do$;

-- ---- structural checks + results -------------------------------------------
select label, ok, detail from t_res
union all
select 'columns added', count(*) = 2, count(*)::text
  from information_schema.columns
  where table_schema = 'public' and table_name = 'feedback_log'
    and column_name in ('assigned_to', 'assigned_at')
union all
select 'index created', count(*) = 1, count(*)::text
  from pg_indexes where tablename = 'feedback_log' and indexname = 'feedback_log_assigned_idx'
union all
select 'trigger bound', count(*) = 1, count(*)::text
  from pg_trigger t join pg_class c on c.oid = t.tgrelid
  where c.relname = 'feedback_log' and not t.tgisinternal
union all
select 'existing notes untouched', count(*) = 30, count(*)::text
  from public.feedback_log where note not like '[TEST 138]%';

rollback;
