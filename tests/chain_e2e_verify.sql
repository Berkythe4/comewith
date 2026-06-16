-- =============================================================================
-- chain_e2e_verify.sql  —  Sprint 4 Phase 2: END-TO-END chain on the REAL DI#2 event
-- (Dance Infusion #2 @ Signal). Walks every link in ONE rolled-back transaction —
-- persists NOTHING. Proves interconnection: venue → contacts surface → generate →
-- one gear task + rider→sound + load-in→booking + contract→booking → venue
-- contractable → delete stays deleted. RAISEs TEST_RESULTS_OK to abort.
-- =============================================================================
do $$
declare
  v_admin uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_event uuid := 'ff2b1917-9a3f-42f8-8b9f-4a1deec2338e';  -- Dance Infusion #2
  v_venue uuid;
  v_sound uuid; v_book uuid; v_vactor uuid;
  v_venue_name text; v_contacts int; v_gear int;
  v_rider_sound boolean; v_loadin_book boolean; v_contract_book boolean;
  v_counterparty_ok boolean; v_gear_after int;
  v_gear_title text := 'Load / test / setup gear (see equipment sheet)';
  v_results jsonb;
begin
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  select venue_id into v_venue from public.events where id = v_event;
  if v_venue is null then raise exception 'PRECHECK FAIL: DI#2 has no venue (the save bug would be unfixed)'; end if;

  -- LINK 1: venue persists + resolves by explicit lookup (the display fix path)
  select name into v_venue_name from public.venues where id = v_venue;

  -- Make DI#2 gear-providing for this walk
  update public.events set cw_providing_gear = true where id = v_event;

  -- LINK 2: build the contact matrix → contacts surface
  select id into v_sound from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_book  from public.actors where deleted_at is null and id <> v_sound order by display_name limit 1;
  insert into public.venue_contacts (venue_id, actor_id, function, is_primary) values (v_venue, v_sound, 'sound', true)
    on conflict (venue_id, actor_id, coalesce(function,'')) do nothing;
  insert into public.venue_contacts (venue_id, actor_id, function, is_primary) values (v_venue, v_book, 'booking', true)
    on conflict (venue_id, actor_id, coalesce(function,'')) do nothing;
  select count(*) into v_contacts from public.v_venue_contacts where venue_id = v_venue;

  -- LINK 5 (set up): venue contractable — link an org actor
  insert into public.actors (display_name, kind) values (v_venue_name||' (org)', 'org') returning id into v_vactor;
  update public.venues set actor_id = v_vactor where id = v_venue;
  select (actor_id = v_vactor) into v_counterparty_ok from public.venues where id = v_venue;

  -- LINK 3: generate the checklist
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_gear from public.tasks where event_id=v_event and title=v_gear_title and deleted_at is null;
  select (ta.actor_id=v_sound) into v_rider_sound from public.tasks t join public.task_assignments ta on ta.task_id=t.id
    where t.event_id=v_event and t.title='Send rider to sound contact' limit 1;
  select (ta.actor_id=v_book) into v_loadin_book from public.tasks t join public.task_assignments ta on ta.task_id=t.id
    where t.event_id=v_event and t.title='Confirm load-in time with venue' limit 1;
  select (ta.actor_id=v_book) into v_contract_book from public.tasks t join public.task_assignments ta on ta.task_id=t.id
    where t.event_id=v_event and t.title='Finalize & sign venue contract' limit 1;

  -- LINK 6: delete a generated task → regenerate → stays deleted
  update public.tasks set deleted_at = now() where event_id=v_event and title=v_gear_title;
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_gear_after from public.tasks where event_id=v_event and title=v_gear_title and deleted_at is null;

  v_results := jsonb_build_object(
    'di2_venue_name', v_venue_name,
    'contacts_surface', v_contacts,
    'one_gear_task', v_gear,
    'rider_to_sound', v_rider_sound,
    'loadin_to_booking', v_loadin_book,
    'contract_to_booking', v_contract_book,
    'venue_contractable', v_counterparty_ok,
    'deleted_stays_deleted', (v_gear_after = 0)
  );

  if v_venue_name is null then raise exception 'ASSERT FAIL: venue name unresolved'; end if;
  if v_contacts < 2 then raise exception 'ASSERT FAIL: contacts did not surface (%)', v_contacts; end if;
  if v_gear <> 1 then raise exception 'ASSERT FAIL: not one gear task (%)', v_gear; end if;
  if v_rider_sound is not true then raise exception 'ASSERT FAIL: rider not -> sound'; end if;
  if v_loadin_book is not true then raise exception 'ASSERT FAIL: load-in not -> booking'; end if;
  if v_contract_book is not true then raise exception 'ASSERT FAIL: contract task not -> booking'; end if;
  if not v_counterparty_ok then raise exception 'ASSERT FAIL: venue not contractable'; end if;
  if v_gear_after <> 0 then raise exception 'ASSERT FAIL: deleted task resurrected on regen'; end if;

  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
