-- =============================================================================
-- 035_chain_fixes.sql  —  Sprint 4 Phase 1: fix the interconnection chain
--
--  A) venues.actor_id — venue-as-counterparty (decision A-a). A venue links to an
--     actor (kind=org) = its contractable legal identity. Nullable; the place stays
--     a venue, the actor is its contract identity.
--  B) v_event_equipment_sheet — the load-in sheet: equipment_usage rows for an event
--     as a clean list (the ONE consolidated gear task points here instead of N tasks).
--  C) generate_day_of_tasks rewritten:
--       - ONE gear task ("Load / test / setup gear (see equipment sheet)") when CW is
--         providing gear — replaces the per-item "Load/test/setup: <item>" ×N noise.
--       - DELETE-RESURRECTION FIX: suppression now matches a title that exists for the
--         event REGARDLESS of deleted_at — so a task the user deleted is NOT re-added on
--         regeneration. (hub task-delete is a soft delete, so the row persists to suppress.)
--         Live duplicates still skipped (idempotent); future-only template edits preserved.
--
-- ADDITIVE ONLY: 1 nullable column, 1 view, CREATE OR REPLACE function. No DROP /
-- destructive ALTER / data deletion.
-- =============================================================================
begin;

-- A) venue ↔ contractable actor
alter table public.venues add column if not exists actor_id uuid references public.actors(id);
comment on column public.venues.actor_id is
  'The venue''s contractable identity (an actor, kind=org). Set when you first contract with the venue; lets the venue be selected as a contract counterparty.';

-- B) load-in sheet view
create or replace view public.v_event_equipment_sheet with (security_invoker = true) as
select eu.event_id, eu.id as usage_id, eu.equipment_id, ei.name, ei.category,
       eu.purpose, eu.start_date, eu.end_date, eu.revenue_attributed
from public.equipment_usage eu
join public.equipment_inventory ei on ei.id = eu.equipment_id;
revoke all on public.v_event_equipment_sheet from anon;

-- C) generator: one gear task, delete-suppression, idempotent
create or replace function public.generate_day_of_tasks(p_event_id uuid)
  returns integer language plpgsql security definer set search_path = public as $$
declare
  v_type text; v_gear boolean; v_venue uuid; v_date date; v_count int := 0;
  r record; v_task uuid; v_actor uuid; v_title text;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;
  select type, coalesce(cw_providing_gear,false), venue_id, event_date
    into v_type, v_gear, v_venue, v_date from public.events where id = p_event_id;

  -- ONE consolidated gear task when CW is providing gear (detail lives on the equipment sheet).
  -- Suppress if a task with this title exists for the event — INCLUDING soft-deleted — so a
  -- deliberately-removed task is never resurrected.
  if v_gear then
    v_title := 'Load / test / setup gear (see equipment sheet)';
    if not exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=v_title) then
      insert into public.tasks(event_id,title,status,source) values (p_event_id,v_title,'todo','template');
      v_count := v_count + 1;
    end if;
  end if;

  -- dj/performer participants → soundcheck (roles[] overlap); delete-suppressed
  for r in select a.display_name nm
             from public.event_participants ep join public.actors a on a.id = ep.actor_id
            where ep.event_id = p_event_id and ep.roles && array['dj','performer','headliner','opener']::text[] loop
    v_title := 'Soundcheck / confirm arrival: '||r.nm;
    if not exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=v_title) then
      insert into public.tasks(event_id,title,status,source) values (p_event_id,v_title,'todo','template');
      v_count := v_count + 1;
    end if;
  end loop;

  -- Template-driven across phases, gear-filtered, ordered; outreach auto-assigns. Delete-suppressed.
  for r in select tt.title, tt.default_offset_days, tt.target_function
             from public.task_templates tt
            where tt.event_type = v_type
              and (tt.gear_applicability = 'both'
                   or tt.gear_applicability = case when v_gear then 'gear' else 'no_gear' end)
            order by tt.phase, tt.sort_order, tt.title loop
    if exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=r.title) then
      continue;
    end if;
    insert into public.tasks(event_id, title, status, source, due_date)
    values (p_event_id, r.title, 'todo', 'template',
            case when r.default_offset_days is not null and v_date is not null then v_date + r.default_offset_days else null end)
    returning id into v_task;
    v_count := v_count + 1;

    if r.target_function is not null then
      v_actor := null;
      if r.target_function like 'venue:%' and v_venue is not null then
        select actor_id into v_actor from public.venue_contacts
         where venue_id = v_venue and function = split_part(r.target_function, ':', 2)
         order by is_primary desc limit 1;
      elsif r.target_function = 'vendor' then
        select ep.actor_id into v_actor from public.event_participants ep
         where ep.event_id = p_event_id and ep.roles && array['vendor']::text[] limit 1;
      end if;
      if v_actor is not null then
        insert into public.task_assignments(task_id, actor_id, role) values (v_task, v_actor, 'doer')
          on conflict (task_id, actor_id, role) do nothing;
      else
        update public.tasks set description = 'Assign a '||r.target_function||' contact — none on file yet (add one on the venue/vendor)'
         where id = v_task;
      end if;
    end if;
  end loop;

  return v_count;
end $$;

commit;

-- =============================================================================
-- DOWN (manual): restore 034 generator body; drop view v_event_equipment_sheet;
-- alter venues drop column actor_id.
-- =============================================================================
