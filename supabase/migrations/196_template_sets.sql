-- ============================================================
-- COME WITH — 196 named task template SETS, not per-event-type lists
--
-- WHY
-- task_templates was a flat list keyed by event_type: a party had exactly one
-- possible checklist, and the only way to change it was to edit the one that
-- every future party would also use. There was no way to keep a v1 around while
-- trying a v2, and no way to have a checklist that isn't about an event type at
-- all.
--
-- Now a SET is a named thing you author on the Templates page ("Event Template
-- v2"), holding ordered steps, and you pick which set to apply when you build an
-- event's checklist.
--
-- event_type is GONE from the model, deliberately (decided with Henry
-- 2026-08-21). Sets are free: any set can be applied to any event. The three
-- existing groups become three named sets, so nothing an event already runs
-- changes, but the type is now just part of a name a human reads rather than a
-- filter the system enforces.
--
-- The column is DROPPED rather than left nullable. A stale column that nothing
-- reads is worse than no column: the next person to touch this will filter on it
-- and get a silently empty list. The set name carries the meaning now.
--
-- events.task_template_set_id records WHICH set an event is running. The
-- calendar gap panel measures "N of M steps missing" against that set — without
-- it, an event deliberately run on v2 would read as permanently missing all of
-- v1's steps. It is written when a workflow actually creates its first task, not
-- when the set is merely picked, so it always describes tasks that exist.
-- ============================================================
begin;

-- ---------- 1. the sets ----------
create table if not exists public.task_template_sets (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  description text,
  active      boolean not null default true,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create unique index if not exists task_template_sets_name_key
  on public.task_template_sets (lower(btrim(name)));

drop trigger if exists set_updated_at on public.task_template_sets;
create trigger set_updated_at before update on public.task_template_sets
  for each row execute function public.handle_updated_at();

alter table public.task_template_sets enable row level security;

-- Same gate the template rows themselves use (042), so the two cannot diverge.
drop policy if exists "Template sets module access" on public.task_template_sets;
create policy "Template sets module access" on public.task_template_sets for all
  using (public.user_can_access_module('templates') or public.can_use_events_module())
  with check (public.user_can_access_module('templates') or public.can_use_events_module());

comment on table public.task_template_sets is
  'A named checklist you can apply to an event ("Party — standard", "Event Template v2"). Deliberately NOT tied to an event type: any set can be applied to any event, and the picker shows them all.';

-- ---------- 2. steps belong to a set ----------
alter table public.task_templates
  add column if not exists set_id uuid references public.task_template_sets(id) on delete cascade;

-- Seed one set per existing event_type group, then link the rows. Named so the
-- old grouping is still legible to a human on the Templates page.
insert into public.task_template_sets (name, description)
select initcap(replace(tt.event_type, '_', ' ')) || ' — standard',
       'Migrated from the ' || tt.event_type || ' checklist (196). Was the only checklist that event type could use.'
  from public.task_templates tt
 where tt.set_id is null
 group by tt.event_type
on conflict do nothing;

update public.task_templates tt
   set set_id = s.id
  from public.task_template_sets s
 where tt.set_id is null
   and s.name = initcap(replace(tt.event_type, '_', ' ')) || ' — standard';

-- Every row must now belong to a set.
alter table public.task_templates alter column set_id set not null;

-- event_type is superseded. Drop the constraint, the old uniqueness rule, and
-- the column — see the header for why this is a drop and not a soft-deprecate.
drop index if exists public.idx_task_templates_unique;
alter table public.task_templates drop constraint if exists task_templates_event_type_check;
alter table public.task_templates drop column if exists event_type;

-- Titles are unique WITHIN a set now. Two different sets are allowed to contain
-- a step of the same name — that is the point of having a v1 and a v2.
create unique index if not exists task_templates_set_title_key
  on public.task_templates (set_id, lower(btrim(title)));

create index if not exists task_templates_set_idx on public.task_templates (set_id);

-- ---------- 3. which set is this event running ----------
alter table public.events
  add column if not exists task_template_set_id uuid references public.task_template_sets(id) on delete set null;

comment on column public.events.task_template_set_id is
  'The template set this event''s checklist was built from. Written when a workflow creates its first task, so it always describes tasks that exist. The calendar gap panel measures missing steps against THIS set.';

-- Backfill: an event whose existing template tasks point into exactly one set is
-- running that set. Events with tasks from more than one set are left null
-- rather than guessed at — a wrong answer here shows up as phantom missing steps.
update public.events e
   set task_template_set_id = x.set_id
  from (
    -- (array_agg)[1] rather than min(): there is no min(uuid) in Postgres.
    select t.event_id, (array_agg(distinct tt.set_id))[1] set_id
      from public.tasks t
      join public.task_templates tt on tt.id = t.template_id
     where t.template_id is not null
     group by t.event_id
    having count(distinct tt.set_id) = 1
  ) x
 where x.event_id = e.id
   and e.task_template_set_id is null;

-- ---------- 4. the planner takes a set ----------
drop function if exists public.plan_event_tasks(uuid);
drop function if exists public.plan_event_tasks(uuid, uuid);

create function public.plan_event_tasks(p_event_id uuid, p_set_id uuid default null)
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
  v_gear boolean; v_venue uuid; v_date date; v_set uuid;
  r record; v_actor uuid; v_name text; v_title text; v_desc text; v_due date;
  v_n integer := 0;
  v_dayof integer := 0;
  v_rank integer;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;

  select coalesce(e.cw_providing_gear, false), e.venue_id, e.event_date, e.task_template_set_id
    into v_gear, v_venue, v_date, v_set
    from public.events e
   where e.id = p_event_id;

  if not found then raise exception 'no such event: %', p_event_id; end if;

  -- An explicit set wins; otherwise plan against whatever this event is already
  -- running. If it is running nothing, only the derived steps are proposed —
  -- the caller is expected to have asked which set to use.
  v_set := coalesce(p_set_id, v_set);

  -- ONE consolidated gear step when CW is providing gear (detail lives on the
  -- equipment sheet). Derived from the event, not from a set, so it is offered
  -- whichever checklist you pick.
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

  if v_set is null then return; end if;

  -- Steps from the chosen set, gear-filtered, in phase order. Outreach steps
  -- carry their suggested assignee out to the caller instead of assigning here —
  -- the planner writes nothing.
  --
  -- Phases are ordered by MEANING, not alphabetically: 'day_of' sorts before
  -- 'planning' as text, which nobody noticed while this was a silent bulk insert.
  for r in
    select tt.id, tt.title, tt.default_offset_days, tt.target_function, tt.phase
      from public.task_templates tt
     where tt.set_id = v_set
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
    -- that the backfill could not reach, and stops a step being offered twice
    -- when two sets happen to share a step name.
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

      if v_actor is null then
        v_desc := 'Assign a ' || r.target_function
               || ' contact — none on file yet (add one on the venue/vendor)';
      end if;
    end if;

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

comment on function public.plan_event_tasks(uuid, uuid) is
  'What generate_day_of_tasks would create for this event from the given set, without creating it. Falls back to events.task_template_set_id when no set is passed. Admin only.';

-- ---------- 5. the generator, still a thin consumer ----------
drop function if exists public.generate_day_of_tasks(uuid);
drop function if exists public.generate_day_of_tasks(uuid, uuid);

create function public.generate_day_of_tasks(p_event_id uuid, p_set_id uuid default null)
returns integer language plpgsql security definer set search_path = public as $$
declare r record; v_task uuid; v_count integer := 0; v_set uuid;
begin
  if not public.is_admin() then raise exception 'admin only'; end if;

  select coalesce(p_set_id, e.task_template_set_id) into v_set
    from public.events e where e.id = p_event_id;

  for r in select * from public.plan_event_tasks(p_event_id, v_set) order by sort_key loop
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

  -- Record what this event is running, but only if it actually got steps from
  -- the set. A run that created nothing must not relabel the event.
  if v_count > 0 and v_set is not null then
    update public.events set task_template_set_id = v_set
     where id = p_event_id and task_template_set_id is distinct from v_set;
  end if;

  return v_count;
end $$;

-- EXECUTE goes to PUBLIC by default and anon inherits it, so `revoke from anon`
-- alone is a no-op — PUBLIC has to go, and authenticated granted back. Both
-- functions still guard themselves with is_admin(). (Learned the hard way in 195.)
revoke all on function public.plan_event_tasks(uuid, uuid)      from public, anon;
revoke all on function public.generate_day_of_tasks(uuid, uuid) from public, anon;
grant execute on function public.plan_event_tasks(uuid, uuid)      to authenticated;
grant execute on function public.generate_day_of_tasks(uuid, uuid) to authenticated;

commit;
