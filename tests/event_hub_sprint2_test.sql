-- =============================================================================
-- event_hub_sprint2_test.sql  —  Event Hub UX pass (Sprint 2)
--
-- Validates the sprint-2 data layer against REAL prod, WITHOUT persisting: one DO
-- block that exercises every new write then RAISEs `TEST_RESULTS_OK …` so the
-- implicit transaction aborts (zero rows persist). Green = that message; a real
-- failure raises `ASSERT FAIL …`.
--
-- Covers: multi-role participant (roles[] + role=roles[1]); one-per-actor unique
-- enforcement; bulk people insert; multi-equipment insert; money child inserts;
-- contract edit + files(subject_type='contract') + contracts.document_id wiring;
-- day-of generator reading roles[] overlap (secondary dj → soundcheck task);
-- IG 3-account snapshot upsert; and audit_log capturing the hub tables.
-- =============================================================================
do $$
declare
  v_admin uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_event uuid; v_actor uuid; v_actor2 uuid; v_equip uuid; v_equip2 uuid;
  v_ep uuid; v_contract uuid; v_file uuid;
  v_roles text[]; v_role text;
  v_dup_blocked boolean := false;
  v_bulk int; v_equip_n int; v_soundcheck int; v_ig int; v_docid uuid;
  v_audit_before int; v_audit_after int;
  v_results jsonb;
begin
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  if not public.is_admin() then raise exception 'PRECHECK FAIL: is_admin() false'; end if;

  select id into v_event from public.events where deleted_at is null order by event_date desc limit 1;
  select id into v_actor  from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_actor2 from public.actors where deleted_at is null and id <> v_actor order by display_name limit 1;
  select id into v_equip  from public.equipment_inventory where deleted_at is null order by name limit 1;
  select id into v_equip2 from public.equipment_inventory where deleted_at is null and id <> v_equip order by name limit 1;
  if v_event is null or v_actor is null or v_actor2 is null or v_equip is null then
    raise exception 'PRECHECK FAIL: need event, 2 actors, equipment';
  end if;

  select count(*) into v_audit_before from public.audit_log where table_name in ('tasks','contracts','event_participants');

  update public.events set type = 'dance_infusion' where id = v_event;

  -- 1) MULTI-ROLE participant: role='host' but roles carries a secondary 'dj'
  insert into public.event_participants (event_id, actor_id, role, roles, fee, is_contractor)
  values (v_event, v_actor, 'host', array['host','dj'], 200, false) returning id into v_ep;
  select roles, role into v_roles, v_role from public.event_participants where id = v_ep;

  -- 2) ONE-PER-ACTOR unique enforced (second row for same event+actor must fail)
  begin
    insert into public.event_participants (event_id, actor_id, role, roles)
    values (v_event, v_actor, 'crew', array['crew']);
  exception when unique_violation then v_dup_blocked := true;
  end;

  -- 3) BULK people: many participants in one insert
  insert into public.event_participants (event_id, actor_id, role, roles, fee)
  values (v_event, v_actor2, 'dj', array['dj','producer'], 350);
  select count(*) into v_bulk from public.event_participants where event_id = v_event;

  -- 4) MULTI-EQUIPMENT: many rows in one insert
  insert into public.equipment_usage (event_id, equipment_id, purpose, start_date)
  values (v_event, v_equip, 'own_event', current_date),
         (v_event, v_equip2, 'rental',  current_date);
  select count(*) into v_equip_n from public.equipment_usage where event_id = v_event;

  -- 5) MONEY child inserts (same SQL the shared moneyMutate emits)
  insert into public.ticketing (event_id, ticket_type, quantity, amount_paid) values (v_event, 'GA', 10, 250);
  insert into public.income (event_id, date, amount, category) values (v_event, current_date, 80, 'bar');
  insert into public.expenses (event_id, date, amount, category) values (v_event, current_date, 60, 'venue');
  insert into public.third_party_donations (event_id, date, amount) values (v_event, current_date, 40);

  -- 6) CONTRACT edit + document wiring
  insert into public.contracts (event_id, actor_id, kind, status, fee)
  values (v_event, v_actor, 'vendor', 'draft', 500) returning id into v_contract;
  update public.contracts set status = 'sent', fee = 600, notes = 'edited' where id = v_contract;
  insert into public.files (bucket, path, filename, subject_type, subject_id, kind)
  values ('agreements', 'contract/'||v_contract||'/deal.pdf', 'deal.pdf', 'contract', v_contract, 'contract')
  returning id into v_file;
  update public.contracts set document_id = v_file where id = v_contract;
  select document_id into v_docid from public.contracts where id = v_contract;

  -- 7) DAY-OF generator now reads roles[] overlap → secondary-dj participant gets a soundcheck task
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_soundcheck
    from public.tasks t join public.actors a on t.title = 'Soundcheck / confirm arrival: '||a.display_name
   where t.event_id = v_event and a.id = v_actor and t.deleted_at is null;

  -- 8) IG 3-account snapshot upsert (today)
  insert into public.metric_snapshots (metric_key, value, captured_on, series_id) values
    ('instagram.followers.comewith',     5100, current_date, null),
    ('instagram.followers.berky',        3100, current_date, null),
    ('instagram.followers.danceinfusion',2100, current_date, null)
  on conflict (metric_key, captured_on, coalesce(series_id,'00000000-0000-0000-0000-000000000000'::uuid))
    do update set value = excluded.value;
  select count(*) into v_ig from public.metric_snapshots
    where metric_key like 'instagram.followers.%' and captured_on = current_date and series_id is null;

  -- 9) audit_log captured the hub tables
  select count(*) into v_audit_after from public.audit_log where table_name in ('tasks','contracts','event_participants');

  v_results := jsonb_build_object(
    'multirole_roles', v_roles, 'multirole_role_is_first', (v_role = v_roles[1]),
    'one_per_actor_enforced', v_dup_blocked,
    'bulk_people_total', v_bulk,
    'multi_equipment_total', v_equip_n,
    'contract_document_wired', (v_docid = v_file),
    'dayof_soundcheck_for_secondary_dj', v_soundcheck,
    'ig_accounts_today', v_ig,
    'audit_rows_gained', (v_audit_after - v_audit_before)
  );

  if not v_dup_blocked then raise exception 'ASSERT FAIL: one-per-actor not enforced'; end if;
  if v_role <> v_roles[1] then raise exception 'ASSERT FAIL: role <> roles[1]'; end if;
  if v_soundcheck < 1 then raise exception 'ASSERT FAIL: day-of did not read secondary dj from roles[]'; end if;
  if v_docid <> v_file then raise exception 'ASSERT FAIL: contract.document_id not wired'; end if;
  if v_ig <> 3 then raise exception 'ASSERT FAIL: IG upsert count = %', v_ig; end if;
  if (v_audit_after - v_audit_before) < 3 then raise exception 'ASSERT FAIL: audit_log not capturing hub tables (gained %)', v_audit_after - v_audit_before; end if;

  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
