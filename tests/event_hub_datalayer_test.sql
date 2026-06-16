-- =============================================================================
-- event_hub_datalayer_test.sql  —  Event Hub sprint (Sprint 1)
--
-- Validates EVERY data-layer operation the Event Hub performs, against the REAL
-- prod schema (columns, CHECK constraints, FKs, unique indexes, the day-of RPC,
-- and v_actor_full) — WITHOUT persisting anything.
--
-- HOW IT STAYS SAFE: the whole thing is a single DO block that performs the
-- operations and then RAISEs an exception (`TEST_RESULTS_OK …`) at the very end.
-- The exception aborts the implicit transaction, so all inserts/updates roll back.
-- A green run reports its findings in the raised message; a real assertion failure
-- raises `ASSERT FAIL …` instead. Either way, zero rows persist.
--
-- The dashboard runs every one of these as an authenticated admin (is_admin()).
-- Here we set request.jwt.claim.sub to a master_admin so is_admin() passes for the
-- SECURITY DEFINER day-of generator; the rest is plain DML against the schema.
--
-- RUN (read-only effect):
--   Management API:  POST /v1/projects/<ref>/database/query  with this file as {query}
--   or psql:         psql "<conn>" -f tests/event_hub_datalayer_test.sql
-- Expect an ERROR whose message starts with "TEST_RESULTS_OK" and contains the JSON.
-- =============================================================================
do $$
declare
  v_admin   uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_event   uuid; v_actor uuid; v_equip uuid; v_orig_type text;
  v_task    uuid; v_ep uuid; v_exp uuid; v_contract uuid; v_file uuid;
  v_ep_role text; v_status text; v_stage text; v_roles text;
  v_rows1 int; v_rows2 int; v_genned int;
  v_results jsonb;
begin
  -- Simulate the signed-in admin (is_admin() resolves via auth.uid()).
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  if not public.is_admin() then raise exception 'PRECHECK FAIL: is_admin() false for %', v_admin; end if;

  select id into v_actor from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_equip from public.equipment_inventory where deleted_at is null limit 1;
  select id, type into v_event, v_orig_type from public.events where deleted_at is null order by event_date desc limit 1;
  if v_admin is null or v_actor is null or v_event is null or v_equip is null then
    raise exception 'PRECHECK FAIL: need an admin, an actor, an event and an equipment row';
  end if;

  -- Give the event a known type so day-of templates fire (rolled back with everything else).
  update public.events set type = 'dance_infusion' where id = v_event;

  -- 1) PEOPLE — add a participant (event_participants)
  insert into public.event_participants (event_id, actor_id, role, fee, is_contractor)
  values (v_event, v_actor, 'dj', 150.00, true) returning id, role into v_ep, v_ep_role;

  -- 2) FEE -> EXPENSE — the manual one-click convert
  insert into public.expenses (event_id, date, amount, category, vendor, description, created_by)
  values (v_event, current_date, 150.00, 'talent', 'TEST actor', 'Participant fee — dj', v_admin)
  returning id into v_exp;

  -- 3) TASKS — task + assignment + inline status update
  insert into public.tasks (event_id, title, source, status)
  values (v_event, 'TEST hub task', 'manual', 'todo') returning id into v_task;
  insert into public.task_assignments (task_id, actor_id, role) values (v_task, v_actor, 'doer');
  update public.tasks set status = 'doing' where id = v_task;
  select status into v_status from public.tasks where id = v_task;

  -- 4) EQUIPMENT — the previously-unwritten event-attach path (equipment_usage)
  insert into public.equipment_usage (event_id, equipment_id, purpose, start_date)
  values (v_event, v_equip, 'own_event', current_date);

  -- 5) STAGE — lifecycle update on events.stage
  update public.events set stage = 'confirmed' where id = v_event;
  select stage into v_stage from public.events where id = v_event;

  -- 6) DAY-OF GENERATOR — idempotency: run twice, template-task count must not grow
  v_genned := public.generate_day_of_tasks(v_event);
  select count(*) into v_rows1 from public.tasks where event_id = v_event and source = 'template' and deleted_at is null;
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_rows2 from public.tasks where event_id = v_event and source = 'template' and deleted_at is null;

  -- 7) CONTRACTS + FILES
  insert into public.contracts (event_id, actor_id, kind, status, fee)
  values (v_event, v_actor, 'vendor', 'draft', 500) returning id into v_contract;
  insert into public.files (bucket, path, filename, subject_type, subject_id, kind)
  values ('agreements', 'event/'||v_event||'/test.pdf', 'test.pdf', 'event', v_event, 'contract')
  returning id into v_file;

  -- 8) v_actor_full returns roles arrayed
  select array_to_string(roles, ',') into v_roles from public.v_actor_full where id = v_actor;

  v_results := jsonb_build_object(
    'participant_added',           v_ep is not null,
    'participant_role',            v_ep_role,
    'fee_expense_created',         v_exp is not null,
    'task_status_after_update',    v_status,
    'task_assignment_ok',          exists(select 1 from public.task_assignments where task_id=v_task and actor_id=v_actor),
    'equipment_usage_ok',          exists(select 1 from public.equipment_usage where event_id=v_event and equipment_id=v_equip),
    'stage_after_update',          v_stage,
    'dayof_first_call_returned',   v_genned,
    'dayof_rows_after_1',          v_rows1,
    'dayof_rows_after_2',          v_rows2,
    'dayof_idempotent',            (v_rows1 = v_rows2 and v_rows1 > 0),
    'contract_created',            v_contract is not null,
    'file_row_created',            v_file is not null,
    'v_actor_full_roles',          v_roles
  );

  -- Hard assertions (raise on failure)
  if v_status <> 'doing'                              then raise exception 'ASSERT FAIL task status = %', v_status; end if;
  if v_stage  <> 'confirmed'                          then raise exception 'ASSERT FAIL stage = %', v_stage; end if;
  if not (v_rows1 = v_rows2 and v_rows1 > 0)          then raise exception 'ASSERT FAIL day-of idempotency %/%', v_rows1, v_rows2; end if;
  if v_ep is null or v_exp is null or v_task is null
     or v_contract is null or v_file is null          then raise exception 'ASSERT FAIL a write returned null'; end if;
  if v_roles is null or v_roles = ''                  then raise exception 'ASSERT FAIL v_actor_full roles empty'; end if;

  -- Success: abort (nothing persists) and surface the findings.
  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
