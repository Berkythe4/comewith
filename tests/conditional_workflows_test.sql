-- =============================================================================
-- conditional_workflows_test.sql  —  Part 3b (rolled-back prod test)
-- Proves: gear flag changes which tasks generate (gear vs no-gear); outreach
-- templates generate at correct offsets and AUTO-ASSIGN via the contact matrix,
-- degrading to unassigned-with-hint when no contact is known; generation is
-- idempotent; and template edits are FUTURE-ONLY (an already-generated task is not
-- rewritten). RAISEs TEST_RESULTS_OK so the txn aborts — zero rows persist.
-- =============================================================================
do $$
declare
  v_admin uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_event uuid; v_venue uuid; v_sound uuid; v_vendor uuid; v_date date := '2026-09-01';
  v_rider_due date; v_rider_assignee uuid; v_rider_count int;
  v_nogear_present boolean; v_gear_present_before boolean; v_gear_present_after boolean;
  v_booking_unassigned boolean; v_booking_hint text;
  v_rider_due_after_edit date; v_team_n int; v_part_n int;
  v_results jsonb;
begin
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  if not public.is_admin() then raise exception 'PRECHECK FAIL: is_admin()'; end if;
  select id into v_sound  from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_vendor from public.actors where deleted_at is null and id <> v_sound order by display_name limit 1;

  -- Fresh, isolated event (rolled back) — type dance_infusion (has outreach seeds), gear OFF
  insert into public.events (name, slug, event_date, status, type, cw_providing_gear)
  values ('TEST 3b event', 'test3b-'||gen_random_uuid(), v_date, 'planning', 'dance_infusion', false)
  returning id into v_event;

  insert into public.venues (name) values ('TEST 3b venue') returning id into v_venue;
  update public.events set venue_id = v_venue where id = v_event;

  -- Matrix: a SOUND contact only (no booking contact) → booking outreach must degrade
  insert into public.venue_contacts (venue_id, actor_id, function, is_primary) values (v_venue, v_sound, 'sound', true);
  -- A vendor participant (for vendor-targeted outreach)
  insert into public.event_participants (event_id, actor_id, role, roles) values (v_event, v_vendor, 'vendor', array['vendor']);

  -- ── Run 1: gear OFF ──────────────────────────────────────────────────────
  perform public.generate_day_of_tasks(v_event);

  v_nogear_present := exists (select 1 from public.tasks where event_id=v_event and title='Confirm house gear / backline with venue' and deleted_at is null);
  v_gear_present_before := exists (select 1 from public.tasks where event_id=v_event and title='Gear breakdown & pack out' and deleted_at is null);

  -- Outreach: rider → sound contact, due = event_date - 14, auto-assigned to v_sound
  select due_date into v_rider_due from public.tasks where event_id=v_event and title='Send rider to sound contact' and deleted_at is null;
  select ta.actor_id into v_rider_assignee
    from public.tasks t join public.task_assignments ta on ta.task_id=t.id
   where t.event_id=v_event and t.title='Send rider to sound contact' limit 1;

  -- Degrade: booking outreach has no booking contact → unassigned + hint
  v_booking_unassigned := not exists (
    select 1 from public.tasks t join public.task_assignments ta on ta.task_id=t.id
     where t.event_id=v_event and t.title='Confirm load-in time with venue');
  select description into v_booking_hint from public.tasks where event_id=v_event and title='Confirm load-in time with venue' and deleted_at is null;

  -- ── Run 2: gear ON, re-run (idempotent + gear set appears) ───────────────
  update public.events set cw_providing_gear = true where id = v_event;
  perform public.generate_day_of_tasks(v_event);
  v_gear_present_after := exists (select 1 from public.tasks where event_id=v_event and title='Gear breakdown & pack out' and deleted_at is null);
  select count(*) into v_rider_count from public.tasks where event_id=v_event and title='Send rider to sound contact' and deleted_at is null;

  -- ── FUTURE-ONLY proof: edit the rider template's offset, re-run; the EXISTING
  --    generated task keeps its original due_date (templates never rewrite live tasks).
  update public.task_templates set default_offset_days = -99 where event_type='dance_infusion' and title='Send rider to sound contact';
  perform public.generate_day_of_tasks(v_event);
  select due_date into v_rider_due_after_edit from public.tasks where event_id=v_event and title='Send rider to sound contact' and deleted_at is null;

  -- Grouped-picker data (the assignee groups read these)
  select count(*) into v_team_n from public.actor_roles where role='team' and active;
  select count(*) into v_part_n from public.event_participants where event_id=v_event;

  v_results := jsonb_build_object(
    'gear_off__nogear_task_present', v_nogear_present,
    'gear_off__gear_task_absent', (not v_gear_present_before),
    'rider_due_is_minus14', (v_rider_due = v_date - 14),
    'rider_autoassigned_to_sound', (v_rider_assignee = v_sound),
    'booking_unassigned_no_contact', v_booking_unassigned,
    'booking_has_hint', (v_booking_hint like '%Assign a venue:booking%'),
    'gear_on__gear_task_present', v_gear_present_after,
    'idempotent_rider_count', v_rider_count,
    'future_only_due_unchanged', (v_rider_due_after_edit = v_date - 14),
    'team_actors', v_team_n, 'event_participants', v_part_n
  );

  if not v_nogear_present then raise exception 'ASSERT FAIL: no-gear task missing on gear-off'; end if;
  if v_gear_present_before then raise exception 'ASSERT FAIL: gear task generated on gear-off'; end if;
  if v_rider_due <> v_date - 14 then raise exception 'ASSERT FAIL: rider offset wrong (%)' , v_rider_due; end if;
  if v_rider_assignee is distinct from v_sound then raise exception 'ASSERT FAIL: rider not auto-assigned to sound contact'; end if;
  if not v_booking_unassigned then raise exception 'ASSERT FAIL: booking task should be unassigned'; end if;
  if v_booking_hint is null or v_booking_hint not like '%Assign a venue:booking%' then raise exception 'ASSERT FAIL: booking hint missing'; end if;
  if not v_gear_present_after then raise exception 'ASSERT FAIL: gear task missing after gear-on'; end if;
  if v_rider_count <> 1 then raise exception 'ASSERT FAIL: not idempotent (rider count %)', v_rider_count; end if;
  if v_rider_due_after_edit <> v_date - 14 then raise exception 'ASSERT FAIL: future-only violated — existing task rewritten'; end if;

  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
