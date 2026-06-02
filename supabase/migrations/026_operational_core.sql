-- =============================================================================
-- 026_operational_core.sql  —  Phase C2: workflow / operational core (additive)
-- Spec §4.5. NOT APPLIED — review before apply. Push held.
--
-- events.stage+owner; files; contracts; tasks+assignments (dormant actor-self via
-- RPC); task_templates + on-demand day-of generator; budget_lines + variance view;
-- touchpoints. Contract/participant fees are NOT auto-posted to actuals (Q3/Q9 —
-- mark paid + reconcile). External-actor access is DORMANT (no login provisioned).
-- =============================================================================

-- ---- Event lifecycle ----
alter table public.events add column if not exists stage text
  check (stage in ('idea','planning','confirmed','live','wrapped','reported'));
alter table public.events add column if not exists owner_actor_id uuid references public.actors(id);

-- ---- files (polymorphic attachments on existing storage buckets) ----
create table public.files (
  id            uuid primary key default gen_random_uuid(),
  bucket        text not null,
  path          text not null,
  filename      text,
  mime          text,
  size          bigint,
  subject_type  text not null check (subject_type in ('event','actor','contract','content_item','task')),
  subject_id    uuid not null,
  kind          text,   -- contract|rider|stage_plot|invoice|receipt|photo|other
  uploaded_by   uuid references public.actors(id),
  created_at    timestamptz not null default now()
);
create index idx_files_subject on public.files(subject_type, subject_id);
alter table public.files enable row level security;
create policy "Admins can manage files" on public.files for all using (public.is_admin());

-- ---- contracts (tied to actor + event + fee + budget) ----
create table public.contracts (
  id            uuid primary key default gen_random_uuid(),
  event_id      uuid references public.events(id) on delete set null,   -- nullable: master/standing
  actor_id      uuid not null references public.actors(id) on delete restrict,
  kind          text not null check (kind in ('event_services','contractor','vendor','rental','sponsor')),
  fee           numeric(10,2),
  status        text not null default 'draft'
                  check (status in ('draft','sent','viewed','signed','countersigned','void')),
  sent_at       timestamptz,
  signed_at     timestamptz,
  paid          boolean not null default false,    -- Q9: mark paid + reconcile manually (no auto-actual)
  paid_at       timestamptz,
  document_id   uuid references public.files(id) on delete set null,
  notes         text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);
create index idx_contracts_event on public.contracts(event_id);
create index idx_contracts_actor on public.contracts(actor_id);
create trigger set_updated_at before update on public.contracts
  for each row execute function public.handle_updated_at();
alter table public.contracts enable row level security;
create policy "Admins can manage contracts" on public.contracts for all using (public.is_admin());
-- DORMANT tier: a counterparty sees ONLY their own contracts (incl. their own fee).
create policy "Actors can read own contracts" on public.contracts
  for select using (actor_id = public.current_actor_id());

-- ---- tasks + assignments ----
create table public.tasks (
  id            uuid primary key default gen_random_uuid(),
  event_id      uuid references public.events(id) on delete set null,
  title         text not null,
  description   text,
  status        text not null default 'todo' check (status in ('todo','doing','blocked','done')),
  priority      text,
  due_date      date,
  effort        integer,
  reward        integer,
  created_by    uuid references public.actors(id),
  source        text not null default 'manual' check (source in ('manual','template','jennifer_import')),
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now(),
  deleted_at    timestamptz
);
create index idx_tasks_event on public.tasks(event_id) where deleted_at is null;
create index idx_tasks_status on public.tasks(status) where deleted_at is null;
create trigger set_updated_at before update on public.tasks
  for each row execute function public.handle_updated_at();
alter table public.tasks enable row level security;
create policy "Admins can manage tasks" on public.tasks for all using (public.is_admin());

create table public.task_assignments (
  id          uuid primary key default gen_random_uuid(),
  task_id     uuid not null references public.tasks(id) on delete cascade,
  actor_id    uuid not null references public.actors(id) on delete cascade,
  role        text not null default 'doer' check (role in ('owner','doer','reviewer')),
  created_at  timestamptz not null default now()
);
create unique index idx_task_assignments_unique on public.task_assignments(task_id, actor_id, role);
create index idx_task_assignments_actor on public.task_assignments(actor_id);
alter table public.task_assignments enable row level security;
create policy "Admins can manage task assignments" on public.task_assignments for all using (public.is_admin());

-- DORMANT tier: an actor sees ONLY tasks assigned to them, and their own assignments.
create policy "Actors can read assigned tasks" on public.tasks
  for select using (exists (
    select 1 from public.task_assignments ta
     where ta.task_id = tasks.id and ta.actor_id = public.current_actor_id()));
create policy "Actors can read own assignments" on public.task_assignments
  for select using (actor_id = public.current_actor_id());

-- Status updates by actors go through a column-scoped RPC (NOT a broad UPDATE
-- policy) so an actor can change ONLY status on a task assigned to them.
create or replace function public.actor_set_task_status(p_task uuid, p_status text)
  returns void language plpgsql security definer set search_path = public as $$
begin
  if p_status not in ('todo','doing','blocked','done') then raise exception 'bad status'; end if;
  if not exists (select 1 from public.task_assignments ta
                  where ta.task_id = p_task and ta.actor_id = public.current_actor_id()) then
    raise exception 'not assigned';
  end if;
  update public.tasks set status = p_status where id = p_task;
end $$;

-- ---- task_templates + on-demand day-of generator ----
create table public.task_templates (
  id                  uuid primary key default gen_random_uuid(),
  event_type          text not null check (event_type in ('party','dance_infusion','production','showcase')),
  title               text not null,
  default_offset_days integer,        -- negative = before event (T-60)
  default_role        text,
  phase               text not null check (phase in ('planning','promo','day_of','wrap')),
  created_at          timestamptz not null default now()
);
alter table public.task_templates enable row level security;
create policy "Admins can manage task templates" on public.task_templates for all using (public.is_admin());

-- a few starter templates (admin-editable later)
insert into public.task_templates (event_type, title, default_offset_days, phase) values
  ('dance_infusion','Book venue',                 -60,'planning'),
  ('dance_infusion','Confirm DJ lineup',          -30,'planning'),
  ('dance_infusion','Sponsor outreach',           -45,'planning'),
  ('dance_infusion','Open ticketing',             -35,'promo'),
  ('dance_infusion','Doors / float ready',          0,'day_of'),
  ('dance_infusion','Post-event reconciliation',   +3,'wrap'),
  ('party','Book venue',                          -45,'planning'),
  ('party','Confirm lineup',                      -21,'planning'),
  ('party','Doors / float ready',                   0,'day_of'),
  ('showcase','Confirm shoot plan',               -14,'planning'),
  ('showcase','Publish content',                   +5,'wrap');

-- Generates day-of tasks ON DEMAND (Q8) from assigned equipment + participants
-- + the fixed day_of templates. Admin-only. Skips titles already present for the
-- event so re-running doesn't duplicate.
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

  for r in select a.display_name nm
             from public.event_participants ep join public.actors a on a.id = ep.actor_id
            where ep.event_id = p_event_id and ep.role in ('dj','performer','headliner','opener') loop
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

-- ---- budget_lines (planned) + variance view (planned vs actual) ----
create table public.budget_lines (
  id             uuid primary key default gen_random_uuid(),
  event_id       uuid references public.events(id) on delete cascade,  -- nullable: type-level / overall
  scope          text not null default 'event' check (scope in ('event','event_type','overall')),
  event_type     text check (event_type in ('party','dance_infusion','production','showcase')),
  category       text not null,        -- venue|production|talent|marketing|... (match public-audit groups)
  planned_amount numeric(10,2) not null default 0,
  direction      text not null default 'expense' check (direction in ('expense','income')),
  contract_id    uuid references public.contracts(id) on delete set null,
  notes          text,
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now()
);
create index idx_budget_lines_event on public.budget_lines(event_id);
create trigger set_updated_at before update on public.budget_lines
  for each row execute function public.handle_updated_at();
alter table public.budget_lines enable row level security;
create policy "Admins can manage budget lines" on public.budget_lines for all using (public.is_admin());

-- Per-event expense variance (best-effort category match to expenses.category).
-- FINANCIAL VIEW — revoked from anon now; MUST also be revoked from authenticated
-- before any external login (see ROADMAP blocker / BUILD_LOG §2).
create or replace view public.v_budget_variance as
select b.event_id,
       b.category,
       b.direction,
       sum(b.planned_amount) as planned,
       coalesce((select sum(e.amount) from public.expenses e
                  where e.event_id = b.event_id and e.deleted_at is null
                    and lower(e.category) = lower(b.category)), 0) as actual_expense,
       sum(b.planned_amount) -
         coalesce((select sum(e.amount) from public.expenses e
                    where e.event_id = b.event_id and e.deleted_at is null
                      and lower(e.category) = lower(b.category)), 0) as variance
  from public.budget_lines b
 where b.direction = 'expense' and b.scope = 'event'
 group by b.event_id, b.category, b.direction;

revoke select on public.v_budget_variance from anon;  -- E1 discipline; financial

-- ---- touchpoints (CRM-lite) ----
create table public.touchpoints (
  id          uuid primary key default gen_random_uuid(),
  actor_id    uuid not null references public.actors(id) on delete cascade,
  event_id    uuid references public.events(id) on delete set null,
  kind        text check (kind in ('email','call','meeting','dm','note')),
  summary     text,
  occurred_at timestamptz not null default now(),
  logged_by   uuid references public.actors(id),
  created_at  timestamptz not null default now()
);
create index idx_touchpoints_actor on public.touchpoints(actor_id);
alter table public.touchpoints enable row level security;
create policy "Admins can manage touchpoints" on public.touchpoints for all using (public.is_admin());

-- Grants: 013 default privileges; admin-only via RLS; dormant actor-self SELECT on
-- tasks/task_assignments/contracts. No anon grants. v_budget_variance anon-revoked.

-- DOWN: drop the 7 tables + 2 functions + view; drop events.stage/owner_actor_id.
