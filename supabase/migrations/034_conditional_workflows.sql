-- =============================================================================
-- 034_conditional_workflows.sql  —  Part 3b: conditional day-of + outreach templates
--
--  A) events.cw_providing_gear — the axis: are WE providing the equipment?
--  B) task_templates additive cols: gear_applicability ('gear'|'no_gear'|'both'),
--     target_function (e.g. 'venue:sound' / 'venue:booking' / 'vendor'), sort_order.
--  C) generate_day_of_tasks rewritten to: filter templates by the gear flag, span ALL
--     phases (due_date = event_date + offset), and AUTO-ASSIGN outreach tasks to the
--     contact-matrix person for the target function (degrading to an unassigned task
--     with a "assign a … contact" hint when none is known). Stays idempotent (skip by
--     title). Existing dynamic equipment/performer tasks preserved.
--  D) Seed outreach + gear/no-gear templates for party + dance_infusion.
--
-- ADDITIVE ONLY: new column, new nullable cols, new unique index, seeds, CREATE OR
-- REPLACE function. No DROP / destructive ALTER / data deletion. Template EDITS apply
-- to FUTURE generation only (generation reads templates at run time; it never rewrites
-- tasks already on an event).
-- =============================================================================
begin;

-- A) gear flag
alter table public.events add column if not exists cw_providing_gear boolean not null default false;
comment on column public.events.cw_providing_gear is
  'True = Come With provides the equipment (drives gear-specific day-of tasks). False = venue/other provides; gear tasks are skipped, house-gear confirmation tasks generate instead.';

-- B) template columns
alter table public.task_templates
  add column if not exists gear_applicability text not null default 'both'
    check (gear_applicability in ('gear','no_gear','both')),
  add column if not exists target_function text,
  add column if not exists sort_order integer not null default 0;
comment on column public.task_templates.target_function is
  'Outreach auto-assign target: ''venue:<function>'' (matches venue_contacts.function, e.g. venue:sound) or ''vendor''. Null = not an outreach task.';

-- one template per (event_type, title) so re-seeding / editor upserts are safe
create unique index if not exists idx_task_templates_unique on public.task_templates(event_type, title);

-- D) seed outreach + gear/no-gear templates for the two active event types
insert into public.task_templates (event_type, title, default_offset_days, phase, gear_applicability, target_function, sort_order)
select et.t, s.title, s.off, s.phase, s.gear, s.tgt, s.ord
from (values ('party'), ('dance_infusion')) as et(t)
cross join (values
  ('Send rider to sound contact',            -14, 'planning', 'both',    'venue:sound',   5),
  ('Confirm house gear / backline with venue', -7, 'planning', 'no_gear', 'venue:sound',   8),
  ('Confirm load-in time with venue',         -7, 'planning', 'both',    'venue:booking', 10),
  ('Confirm vendor arrival window',           -3, 'planning', 'both',    'vendor',        15),
  ('Confirm final headcount with venue',      -2, 'planning', 'both',    'venue:booking', 20),
  ('Gear breakdown & pack out',                0, 'day_of',   'gear',    null,            50),
  ('Return / store equipment',                 1, 'wrap',     'gear',    null,            60)
) as s(title, off, phase, gear, tgt, ord)
on conflict (event_type, title) do nothing;

-- C) conditional + outreach generator
create or replace function public.generate_day_of_tasks(p_event_id uuid)
  returns integer language plpgsql security definer set search_path = public as $$
declare
  v_type text; v_gear boolean; v_venue uuid; v_date date; v_count int := 0;
  r record; v_task uuid; v_actor uuid; v_title text;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;
  select type, coalesce(cw_providing_gear,false), venue_id, event_date
    into v_type, v_gear, v_venue, v_date from public.events where id = p_event_id;

  -- Dynamic: assigned equipment → load/test/setup task (idempotent by title)
  for r in select coalesce(ei.name,'equipment') nm
             from public.equipment_usage eu join public.equipment_inventory ei on ei.id = eu.equipment_id
            where eu.event_id = p_event_id loop
    v_title := 'Load / test / setup: '||r.nm;
    if not exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=v_title and t.deleted_at is null) then
      insert into public.tasks(event_id,title,status,source) values (p_event_id,v_title,'todo','template');
      v_count := v_count + 1;
    end if;
  end loop;

  -- Dynamic: dj/performer participants → soundcheck (roles[] overlap)
  for r in select a.display_name nm
             from public.event_participants ep join public.actors a on a.id = ep.actor_id
            where ep.event_id = p_event_id and ep.roles && array['dj','performer','headliner','opener']::text[] loop
    v_title := 'Soundcheck / confirm arrival: '||r.nm;
    if not exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=v_title and t.deleted_at is null) then
      insert into public.tasks(event_id,title,status,source) values (p_event_id,v_title,'todo','template');
      v_count := v_count + 1;
    end if;
  end loop;

  -- Template-driven across ALL phases, gear-filtered, ordered; outreach auto-assigns
  for r in select tt.title, tt.default_offset_days, tt.target_function
             from public.task_templates tt
            where tt.event_type = v_type
              and (tt.gear_applicability = 'both'
                   or tt.gear_applicability = case when v_gear then 'gear' else 'no_gear' end)
            order by tt.phase, tt.sort_order, tt.title loop
    if exists (select 1 from public.tasks t where t.event_id=p_event_id and t.title=r.title and t.deleted_at is null) then
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
-- DOWN (manual): restore the 032-era generator body; drop seeded outreach templates
-- by (event_type,title); drop index idx_task_templates_unique; alter task_templates
-- drop column sort_order, target_function, gear_applicability; alter events drop
-- column cw_providing_gear.
-- =============================================================================
