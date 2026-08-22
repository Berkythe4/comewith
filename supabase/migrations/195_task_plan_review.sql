-- ============================================================
-- COME WITH — 195 plan/apply split for the event workflow generator
--
-- WHY
-- Applying a template was a single server-side shot: generate_day_of_tasks()
-- decided what to create and created it in the same transaction, so the user
-- first saw the tasks after they existed. The dashboard now walks the workflow
-- one step at a time and lets each step be edited or skipped before it is
-- created, which needs the DECISION separated from the WRITE.
--
--   plan_event_tasks(event)   -> what WOULD be created, writes nothing
--   generate_day_of_tasks(ev) -> loops the plan and inserts (unchanged signature)
--
-- The generator is now a thin consumer of the planner, so the review flow and
-- the bulk "Fill in N" button cannot drift apart. That mattered: the calendar
-- gap panel already re-implements the generator's filter client-side to promise
-- an accurate count, and a third copy of the rules was one copy too many.
--
-- tasks.template_id
-- Title was the ONLY link between a template and the task it generated —
-- suppression (035) and the gap panel's "N of M steps missing" both matched on
-- lower(trim(title)). That was survivable while titles were machine-written.
-- The review flow deliberately invites renaming ("Confirm vendor arrival window"
-- -> "Check with Sal re: 6pm drop"), and under title-matching a rename means the
-- step reads as missing forever AND the next run re-creates the original
-- alongside it. So the link becomes an id, with the title match kept as a
-- fallback for rows created before this migration (and backfilled below).
--
-- Suppression still deliberately counts SOFT-DELETED tasks (035): a step you
-- deleted on purpose must not come back. Skipping a step in the review flow is
-- NOT a delete — nothing is written, so it stays "missing" and can be generated
-- later. That is the intent: skip is for this event, removing it for good is an
-- edit on the Templates page.
-- ============================================================
begin;

-- ---------- 1. the template -> task link ----------
alter table public.tasks
  add column if not exists template_id uuid references public.task_templates(id) on delete set null;

create index if not exists tasks_template_idx on public.tasks(template_id) where template_id is not null;

comment on column public.tasks.template_id is
  'The task_templates row this task was generated from. Survives a retitle, which title-matching did not. Null for manual tasks and for the gear/soundcheck steps, which are derived from the event rather than from a template row.';

-- Backfill: link the tasks the old title-matching generator created. Scoped by
-- event type as well as title, so two event types sharing a step name cannot
-- cross-link. Only source='template' rows — a manual task that happens to share
-- a template's title was never a generated step and must not start behaving
-- like one.
update public.tasks t
   set template_id = tt.id
  from public.task_templates tt, public.events e
 where t.event_id = e.id
   and tt.event_type = e.type
   and lower(btrim(t.title)) = lower(btrim(tt.title))
   and t.source = 'template'
   and t.template_id is null;

-- ---------- 2. the planner ----------
-- Returns the rows generate_day_of_tasks would create, in the order it would
-- create them, and writes nothing. Same admin gate as the generator: this reads
-- the venue contact graph and the participant list, so it is not public data.
drop function if exists public.plan_event_tasks(uuid);

create function public.plan_event_tasks(p_event_id uuid)
returns table (
  kind                 text,
  template_id          uuid,
  title                text,
  description          text,
  due_date             date,
  phase                text,
  suggested_actor_id   uuid,
  suggested_actor_name text,
  sort_key             integer
)
language plpgsql security definer set search_path = public as $$
declare
  v_type text; v_gear boolean; v_venue uuid; v_date date;
  r record; v_actor uuid; v_name text; v_title text; v_desc text; v_due date;
  v_n integer := 0;      -- running counter within the template loop
  v_dayof integer := 0;  -- running counter for the derived day-of steps
  v_rank integer;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;

  select e.type, coalesce(e.cw_providing_gear, false), e.venue_id, e.event_date
    into v_type, v_gear, v_venue, v_date
    from public.events e
   where e.id = p_event_id;

  if not found then raise exception 'no such event: %', p_event_id; end if;

  -- ONE consolidated gear step when CW is providing gear (detail lives on the
  -- equipment sheet). Not a template row, so it stays title-suppressed.
  if v_gear then
    v_title := 'Load / test / setup gear (see equipment sheet)';
    if not exists (
      select 1 from public.tasks t
       where t.event_id = p_event_id
         and lower(btrim(t.title)) = lower(btrim(v_title))
    ) then
      v_dayof := v_dayof + 1;
      return query select 'gear'::text, null::uuid, v_title, null::text, null::date,
                          'day_of'::text, null::uuid, null::text, 3000 + v_dayof;
    end if;
  end if;

  -- One soundcheck step per performing participant (roles[] overlap).
  for r in
    select a.display_name nm
      from public.event_participants ep
      join public.actors a on a.id = ep.actor_id
     where ep.event_id = p_event_id
       and ep.roles && array['dj','performer','headliner','opener']::text[]
     order by a.display_name
  loop
    v_title := 'Soundcheck / confirm arrival: ' || r.nm;
    if not exists (
      select 1 from public.tasks t
       where t.event_id = p_event_id
         and lower(btrim(t.title)) = lower(btrim(v_title))
    ) then
      v_dayof := v_dayof + 1;
      return query select 'soundcheck'::text, null::uuid, v_title, null::text, null::date,
                          'day_of'::text, null::uuid, null::text, 3000 + v_dayof;
    end if;
  end loop;

  -- Template-driven across phases, gear-filtered, ordered. Outreach steps carry
  -- their suggested assignee out to the caller instead of assigning here — the
  -- planner writes nothing, so the generator (or the review modal) does it.
  --
  -- Phases are ordered by MEANING, not alphabetically. `order by tt.phase` put
  -- day_of before planning (d < p), which nobody noticed while this was a silent
  -- bulk insert — the board sorts by due date afterwards either way. The review
  -- modal walks these rows in order and shows the user one at a time, so the
  -- sequence is now something a person reads: plan it, promote it, run it, wrap it.
  for r in
    select tt.id, tt.title, tt.default_offset_days, tt.target_function, tt.phase
      from public.task_templates tt
     where tt.event_type = v_type
       and (tt.gear_applicability = 'both'
            or tt.gear_applicability = case when v_gear then 'gear' else 'no_gear' end)
     order by case tt.phase
                when 'planning' then 1
                when 'promo'    then 2
                when 'day_of'   then 3
                when 'wrap'     then 4
                else 5
              end,
              tt.sort_order, tt.title
  loop
    -- Matched by id first; the title fallback covers rows generated before 195
    -- that the backfill could not reach (event type changed since, say).
    if exists (
      select 1 from public.tasks t
       where t.event_id = p_event_id
         and (t.template_id = r.id or lower(btrim(t.title)) = lower(btrim(r.title)))
    ) then
      continue;
    end if;

    v_due := case
               when r.default_offset_days is not null and v_date is not null
                 then v_date + r.default_offset_days
               else null
             end;

    v_actor := null; v_name := null; v_desc := null;

    if r.target_function is not null then
      if r.target_function like 'venue:%' and v_venue is not null then
        select vc.actor_id, a.display_name
          into v_actor, v_name
          from public.venue_contacts vc
          join public.actors a on a.id = vc.actor_id
         where vc.venue_id = v_venue
           and vc.function = split_part(r.target_function, ':', 2)
         order by vc.is_primary desc
         limit 1;
      elsif r.target_function = 'vendor' then
        select ep.actor_id, a.display_name
          into v_actor, v_name
          from public.event_participants ep
          join public.actors a on a.id = ep.actor_id
         where ep.event_id = p_event_id
           and ep.roles && array['vendor']::text[]
         limit 1;
      end if;

      -- No contact on file: the step still gets made, carrying the reason it
      -- has nobody on it. Pre-195 this was written after the insert; now it
      -- rides the plan, so the review modal shows it before the task exists.
      if v_actor is null then
        v_desc := 'Assign a ' || r.target_function
               || ' contact — none on file yet (add one on the venue/vendor)';
      end if;
    end if;

    -- sort_key puts every step on one timeline: phase band * 1000, the derived
    -- gear/soundcheck steps at the head of the day_of band (you set gear up
    -- before doors), template steps from x100 within their own band. Callers
    -- order by it rather than trusting the order rows happen to come back in.
    v_rank := case r.phase
                when 'planning' then 1
                when 'promo'    then 2
                when 'day_of'   then 3
                when 'wrap'     then 4
                else 5
              end;

    v_n := v_n + 1;
    return query select 'template'::text, r.id, r.title, v_desc, v_due,
                        r.phase, v_actor, v_name, v_rank * 1000 + 100 + v_n;
  end loop;

  return;
end $$;

comment on function public.plan_event_tasks(uuid) is
  'What generate_day_of_tasks would create for this event, without creating it. Admin only. The dashboard walks these rows one modal at a time so each step can be edited or skipped before it is written.';

-- ---------- 3. the generator, now a consumer of the plan ----------
-- Same signature, same return (rows created), same delete-suppression. The
-- decision logic no longer lives here at all.
create or replace function public.generate_day_of_tasks(p_event_id uuid)
returns integer language plpgsql security definer set search_path = public as $$
declare r record; v_task uuid; v_count integer := 0;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;

  for r in select * from public.plan_event_tasks(p_event_id) order by sort_key loop
    insert into public.tasks(event_id, title, description, status, source, due_date, template_id)
    values (p_event_id, r.title, r.description, 'todo', 'template', r.due_date, r.template_id)
    returning id into v_task;

    v_count := v_count + 1;

    if r.suggested_actor_id is not null then
      insert into public.task_assignments(task_id, actor_id, role)
      values (v_task, r.suggested_actor_id, 'doer')
      on conflict (task_id, actor_id, role) do nothing;
    end if;
  end loop;

  return v_count;
end $$;

-- Both guard themselves with is_admin(), but lock the ACL down anyway — the 183
-- audit again. NOTE the order matters and `from anon` alone is a no-op: Postgres
-- grants EXECUTE on a new function to PUBLIC, and anon inherits that, so
-- `revoke ... from anon` leaves the function wide open (verified — it reported
-- anon could still execute). PUBLIC has to be revoked, which also takes EXECUTE
-- away from authenticated, so that one is granted back explicitly.
revoke all on function public.plan_event_tasks(uuid)      from public, anon;
revoke all on function public.generate_day_of_tasks(uuid) from public, anon;
grant execute on function public.plan_event_tasks(uuid)      to authenticated;
grant execute on function public.generate_day_of_tasks(uuid) to authenticated;

commit;
