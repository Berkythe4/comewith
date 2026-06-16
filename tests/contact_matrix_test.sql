-- =============================================================================
-- contact_matrix_test.sql  —  Venue module Part 3a (GATE before 3b)
-- Rolled-back prod test: every 3a write/read, then RAISE TEST_RESULTS_OK so the
-- transaction aborts (zero rows persist). Green = that message in the error.
-- =============================================================================
do $$
declare
  v_admin uuid := (select id from public.profiles where role in ('master_admin','sub_admin') order by role limit 1);
  v_venue uuid; v_event uuid; v_a1 uuid; v_a2 uuid; v_vendor uuid;
  v_vc_count int; v_a1_fn text; v_a1_primary boolean; v_a1_last date;
  v_evd date; v_vendc int; v_dup_blocked boolean := false; v_cap int;
  v_results jsonb;
begin
  perform set_config('request.jwt.claim.sub', v_admin::text, true);
  select id into v_a1 from public.actors where deleted_at is null order by display_name limit 1;
  select id into v_a2 from public.actors where deleted_at is null and id <> v_a1 order by display_name limit 1;
  select id, event_date into v_event, v_evd from public.events where deleted_at is null order by event_date desc limit 1;
  if v_a1 is null or v_a2 is null or v_event is null then raise exception 'PRECHECK FAIL'; end if;

  -- 1) Venue CRUD
  insert into public.venues (name, city, state, capacity) values ('TEST Venue', 'Brooklyn', 'NY', 200) returning id into v_venue;
  update public.venues set capacity = 250 where id = v_venue;
  select capacity into v_cap from public.venues where id = v_venue;

  -- 2) Venue contacts (actors) + role tag
  insert into public.venue_contacts (venue_id, actor_id, function, is_primary) values (v_venue, v_a1, 'booking', true);
  insert into public.venue_contacts (venue_id, actor_id, function) values (v_venue, v_a2, 'sound');
  insert into public.actor_roles (actor_id, role) values (v_a1, 'venue_contact') on conflict (actor_id, role) do nothing;

  -- one-per-(venue,actor,function) enforced
  begin
    insert into public.venue_contacts (venue_id, actor_id, function) values (v_venue, v_a1, 'booking');
  exception when unique_violation then v_dup_blocked := true; end;

  -- 3) Set venue on event + make a1 participate so last_event_with populates
  update public.events set venue_id = v_venue where id = v_event;
  insert into public.event_participants (event_id, actor_id, role, roles) values (v_event, v_a1, 'venue_contact', array['venue_contact'])
    on conflict (event_id, actor_id) do nothing;

  -- 4) Read the matrix view (lookup source)
  select count(*) into v_vc_count from public.v_venue_contacts where venue_id = v_venue;
  select function, is_primary, last_event_with into v_a1_fn, v_a1_primary, v_a1_last
    from public.v_venue_contacts where venue_id = v_venue and actor_id = v_a1;

  -- 5) Vendor + vendor contact
  insert into public.actors (display_name, kind) values ('TEST Vendor Co', 'org') returning id into v_vendor;
  insert into public.actor_roles (actor_id, role) values (v_vendor, 'vendor');
  insert into public.vendor_contacts (vendor_actor_id, contact_actor_id, function, is_primary) values (v_vendor, v_a2, 'booking', true);
  select count(*) into v_vendc from public.v_vendor_contacts where vendor_actor_id = v_vendor;

  v_results := jsonb_build_object(
    'venue_created_updated_cap', v_cap,
    'venue_contacts_count', v_vc_count,
    'a1_function', v_a1_fn,
    'a1_is_primary', v_a1_primary,
    'a1_last_event_with', v_a1_last,
    'event_venue_set', (select venue_id = v_venue from public.events where id = v_event),
    'one_per_function_enforced', v_dup_blocked,
    'vendor_contacts_count', v_vendc
  );

  if v_cap <> 250 then raise exception 'ASSERT FAIL venue update'; end if;
  if v_vc_count <> 2 then raise exception 'ASSERT FAIL venue_contacts count = %', v_vc_count; end if;
  if v_a1_fn <> 'booking' or not v_a1_primary then raise exception 'ASSERT FAIL a1 function/primary'; end if;
  if v_a1_last is null then raise exception 'ASSERT FAIL last_event_with not populated (recency seam)'; end if;
  if not v_dup_blocked then raise exception 'ASSERT FAIL one-per-function not enforced'; end if;
  if v_vendc <> 1 then raise exception 'ASSERT FAIL vendor_contacts count = %', v_vendc; end if;

  raise exception 'TEST_RESULTS_OK %', v_results::text;
end $$;
