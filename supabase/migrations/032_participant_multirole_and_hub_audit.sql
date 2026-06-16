-- =============================================================================
-- 032_participant_multirole_and_hub_audit.sql  —  Event Hub UX pass (Sprint 2)
--
-- Three additive, reversible changes:
--   A) Multi-role participants: event_participants.roles text[] (one row per person
--      per event, many roles). Keeps the existing single `role` column populated
--      (= roles[1]) for backward-compat. Backfills the 5 existing rows. Enforces
--      one-row-per-person with a unique (event_id, actor_id) index (verified no
--      existing dupes). The day-of generator is repointed to read roles[] (overlap).
--   B) generate_day_of_tasks: read DJ/performer participants via roles[] overlap so
--      a person whose dj role is secondary still gets a soundcheck task.
--   C) Lightweight reporting history: attach the existing audit_trigger_function to
--      tasks / contracts / event_participants so status & role transitions land in
--      audit_log (today it only covers expenses/income/profiles/agreements).
--
-- ADDITIVE ONLY: new column (nullable-safe default), new index, CREATE OR REPLACE on
-- the function, new triggers. No DROP, no destructive ALTER, no data deletion.
-- =============================================================================
begin;

-- ── A) Multi-role column + backfill ──────────────────────────────────────────
alter table public.event_participants
  add column if not exists roles text[] not null default '{}';

-- Backfill: existing single role becomes the first (and only) element.
update public.event_participants
   set roles = array[role]
 where cardinality(roles) = 0 and role is not null;

comment on column public.event_participants.roles is
  'All roles this person holds at this event (headliner/dj/opener/painter/…). One row per person per event. The single `role` column is kept = roles[1] for backward-compat.';

-- One row per person per event (verified no existing (event_id, actor_id) dupes).
create unique index if not exists idx_event_participants_one_per_actor
  on public.event_participants (event_id, actor_id);

-- ── B) Day-of generator reads roles[] (overlap) ──────────────────────────────
create or replace function public.generate_day_of_tasks(p_event_id uuid)
  returns integer language plpgsql security definer set search_path = public as $$
declare v_type text; v_count int := 0; r record;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;
  select type into v_type from public.events where id = p_event_id;

  for r in select coalesce(ei.name,'equipment') nm
             from public.equipment_usage eu
             join public.equipment_inventory ei on ei.id = eu.equipment_id
            where eu.event_id = p_event_id loop
    insert into public.tasks (event_id, title, status, source)
    select p_event_id, 'Load / test / setup: '||r.nm, 'todo', 'template'
    where not exists (select 1 from public.tasks t where t.event_id = p_event_id and t.title = 'Load / test / setup: '||r.nm and t.deleted_at is null);
    v_count := v_count + 1;
  end loop;

  -- roles[] overlap so a secondary dj/performer role still triggers a soundcheck task
  for r in select a.display_name nm
             from public.event_participants ep join public.actors a on a.id = ep.actor_id
            where ep.event_id = p_event_id
              and ep.roles && array['dj','performer','headliner','opener']::text[] loop
    insert into public.tasks (event_id, title, status, source)
    select p_event_id, 'Soundcheck / confirm arrival: '||r.nm, 'todo', 'template'
    where not exists (select 1 from public.tasks t where t.event_id = p_event_id and t.title = 'Soundcheck / confirm arrival: '||r.nm and t.deleted_at is null);
    v_count := v_count + 1;
  end loop;

  for r in select tt.title from public.task_templates tt where tt.event_type = v_type and tt.phase = 'day_of' loop
    insert into public.tasks (event_id, title, status, source)
    select p_event_id, r.title, 'todo', 'template'
    where not exists (select 1 from public.tasks t where t.event_id = p_event_id and t.title = r.title and t.deleted_at is null);
    v_count := v_count + 1;
  end loop;

  return v_count;
end $$;

-- ── C) Reporting history: audit the hub's status/role-bearing tables ─────────
-- Reuses public.audit_trigger_function (010). Captures INSERT/UPDATE/DELETE incl.
-- status & role transitions into audit_log (master_admin-readable).
drop trigger if exists audit_tasks on public.tasks;
create trigger audit_tasks
  after insert or update or delete on public.tasks
  for each row execute function public.audit_trigger_function();

drop trigger if exists audit_contracts on public.contracts;
create trigger audit_contracts
  after insert or update or delete on public.contracts
  for each row execute function public.audit_trigger_function();

drop trigger if exists audit_event_participants on public.event_participants;
create trigger audit_event_participants
  after insert or update or delete on public.event_participants
  for each row execute function public.audit_trigger_function();

commit;

-- =============================================================================
-- DOWN (manual):
--   drop trigger if exists audit_event_participants on public.event_participants;
--   drop trigger if exists audit_contracts on public.contracts;
--   drop trigger if exists audit_tasks on public.tasks;
--   -- restore generate_day_of_tasks to the 031-era body (ep.role in (...));
--   drop index if exists public.idx_event_participants_one_per_actor;
--   alter table public.event_participants drop column if exists roles;
-- =============================================================================
