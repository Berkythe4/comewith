with ev as (select id from events where slug=:SLUG)
select json_build_object(
 'events', (select json_agg(e) from events e where deleted_at is null),
 'event_control', (select json_agg(x) from event_control x where event_id=(select id from ev)),
 'event_ticket_tiers', (select json_agg(x) from event_ticket_tiers x where event_id=(select id from ev)),
 'ticketing', (select json_agg(x) from ticketing x where event_id=(select id from ev)),
 'third_party_donations', (select json_agg(x) from third_party_donations x where event_id=(select id from ev)),
 'sponsorships', (select json_agg(to_jsonb(s) || jsonb_build_object('actor', jsonb_build_object('id',a.id,'display_name',a.display_name,'email',a.email))) from sponsorships s left join actors a on a.id=s.actor_id where s.event_id=(select id from ev)),
 'guest_event_attendance', (select json_agg(x) from guest_event_attendance x where event_id=(select id from ev)),
 'event_drink_allocations', (select json_agg(x) from event_drink_allocations x where event_id=(select id from ev)),
 'tasks', (select json_agg(x) from tasks x where event_id=(select id from ev) and deleted_at is null),
 'budget_lines', (select json_agg(x) from budget_lines x where event_id=(select id from ev) and deleted_at is null),
 'expenses', (select json_agg(x) from expenses x where event_id=(select id from ev) and deleted_at is null),
 'income', '[]'::json, 'event_exec_updates', '[]'::json,
 'owner_advances', (select json_agg(x) from owner_advances x where event_id=(select id from ev)),
 'event_participants', (select json_agg(to_jsonb(p) || jsonb_build_object('actor', jsonb_build_object('id',a.id,'display_name',a.display_name,'email',a.email))) from event_participants p join actors a on a.id=p.actor_id where p.event_id=(select id from ev)),
 'event_rsvps', (select json_agg(x) from event_rsvps x where event_id=(select id from ev)),
 'venues', (select json_agg(x) from venues x where id=(select venue_id from events where id=(select id from ev))),
 'guests', (select json_agg(json_build_object('id',id,'full_name',full_name,'email',email,'actor_id',actor_id)) from guests where deleted_at is null),
 'profiles', json_build_array(json_build_object('id','u1','full_name','Keith','email','k@x','role','master_admin','deleted_at',null))
) r;
