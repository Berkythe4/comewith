-- =============================================================================
-- chain_fixes_test.sql  —  Sprint 4 Phase 1 (rolled-back prod test, persists nothing)
-- Proves: ONE consolidated gear task (not N); outreach auto-assign with a venue
-- present; delete-then-regenerate does NOT resurrect; venue-as-counterparty link;
-- equipment-sheet view. RAISEs TEST_RESULTS_OK so the txn aborts.
-- =============================================================================
do $$
declare
  v_admin uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_event uuid; v_venue uuid; v_sound uuid; v_equip uuid; v_vactor uuid;
  v_gear_title text := 'Load / test / setup gear (see equipment sheet)';
  v_gear_live int; v_gear_total int; v_per_item int; v_rider_sound boolean;
  v_gear_live_after int; v_gear_total_after int; v_sheet int; v_link_ok boolean;
  v_results jsonb;
begin
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  select id into v_sound from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_equip from public.equipment_inventory where deleted_at is null order by name limit 1;

  insert into public.events (name, slug, event_date, status, type, cw_providing_gear)
  values ('TEST chain', 'testchain-'||gen_random_uuid(), '2026-10-01', 'planning', 'dance_infusion', true)
  returning id into v_event;
  insert into public.venues (name) values ('TEST chain venue') returning id into v_venue;
  update public.events set venue_id = v_venue where id = v_event;
  insert into public.venue_contacts (venue_id, actor_id, function, is_primary) values (v_venue, v_sound, 'sound', true);
  insert into public.equipment_usage (event_id, equipment_id, purpose, start_date) values (v_event, v_equip, 'own_event', current_date);

  -- Generate (gear ON)
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_gear_live  from public.tasks where event_id=v_event and title=v_gear_title and deleted_at is null;
  select count(*) into v_per_item   from public.tasks where event_id=v_event and title like 'Load / test / setup: %' and deleted_at is null;
  select (ta.actor_id = v_sound) into v_rider_sound
    from public.tasks t join public.task_assignments ta on ta.task_id=t.id
   where t.event_id=v_event and t.title='Send rider to sound contact' limit 1;

  -- Delete-then-regenerate: soft-delete the gear task, regenerate, must NOT resurrect
  update public.tasks set deleted_at = now() where event_id=v_event and title=v_gear_title;
  perform public.generate_day_of_tasks(v_event);
  select count(*) into v_gear_live_after  from public.tasks where event_id=v_event and title=v_gear_title and deleted_at is null;
  select count(*) into v_gear_total_after from public.tasks where event_id=v_event and title=v_gear_title;

  -- Venue-as-counterparty link
  insert into public.actors (display_name, kind) values ('TEST chain venue', 'org') returning id into v_vactor;
  update public.venues set actor_id = v_vactor where id = v_venue;
  select (actor_id = v_vactor) into v_link_ok from public.venues where id = v_venue;

  -- Equipment sheet view
  select count(*) into v_sheet from public.v_event_equipment_sheet where event_id = v_event;

  v_results := jsonb_build_object(
    'one_gear_task', v_gear_live, 'no_per_item_tasks', v_per_item,
    'rider_autoassigned_sound', v_rider_sound,
    'deleted_stays_deleted_live', v_gear_live_after, 'gear_rows_total_after_regen', v_gear_total_after,
    'venue_actor_linked', v_link_ok, 'equipment_sheet_rows', v_sheet
  );

  if v_gear_live <> 1 then raise exception 'ASSERT FAIL: expected 1 gear task, got %', v_gear_live; end if;
  if v_per_item <> 0 then raise exception 'ASSERT FAIL: per-item gear tasks still generated (%)', v_per_item; end if;
  if v_rider_sound is not true then raise exception 'ASSERT FAIL: rider not auto-assigned to sound contact'; end if;
  if v_gear_live_after <> 0 then raise exception 'ASSERT FAIL: deleted gear task was RESURRECTED'; end if;
  if v_gear_total_after <> 1 then raise exception 'ASSERT FAIL: regen created a duplicate gear row (total %)', v_gear_total_after; end if;
  if not v_link_ok then raise exception 'ASSERT FAIL: venue.actor_id link'; end if;
  if v_sheet < 1 then raise exception 'ASSERT FAIL: equipment sheet empty'; end if;

  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
