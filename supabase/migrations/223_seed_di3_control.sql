-- =============================================================================
-- 223_seed_di3_control.sql — Dance Infusion #3 into the Event Control Center
-- =============================================================================
-- Data only, on top of 222. Every write is keyed to the DI #3 event row
-- (slug dance-infusion-3) and asserts what it touched, so a re-run or a wrong
-- target fails loudly instead of doubling anything.
--
-- What is TBD stays TBD - none of it is entered as 0:
--   drink tickets per tier / per Champion, the count locked with Signal, what a
--   drink ticket is worth toward the bar minimum, the start/end time, the name
--   of the $15 portion of the $40 ticket, and three cost lines. Each one is a
--   blocker task (status 'blocked') so it shows on the Command screen AND on
--   Calendar & Tasks until someone answers it.
--
-- events.doors_time already reads 15:00 on this row (it came from DI #1). It is
-- NOT cleared here: the time blocker says it is unconfirmed, and clearing a
-- field the public site may read is not a seed's call to make.
-- =============================================================================

begin;

do $$
declare
  ev      uuid;
  sp_pd   uuid;  sp_tb uuid;  sp_sa uuid;
  g_alice uuid;
  n       int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 (slug dance-infusion-3) not found'; end if;
  if exists (select 1 from public.event_ticket_tiers where event_id = ev) then
    raise exception 'DI #3 already has ticket tiers - seed already applied?';
  end if;

  -- the event: bar minimum (was null)
  update public.events set bar_minimum = 6000 where id = ev and bar_minimum is null;

  insert into public.event_control (event_id, fundraising, beneficiary, mission_target_pct,
                                    drink_tickets_locked, drink_ticket_value, notes)
  values (ev, true, 'National MS Society', 60, null, null,
          'Signal NYC, 175 Morgan Ave. Bar minimum $6,000. Locked drink-ticket count and per-ticket value TBD with Signal.');

  -- tiers: the production share is price - to_mission, derived, never stored
  insert into public.event_ticket_tiers (event_id, label, price, to_mission, production_label, drink_tickets, sort_order) values
    (ev, '$25 ticket', 25, 25, null,                 null, 1),
    (ev, '$40 ticket', 40, 25, 'Production support', null, 2);

  -- sponsors: three Champions, $500 cash each, pending, drink tickets TBD
  insert into public.sponsorships (event_id, actor_id, tier, cash_amount, entry_tickets, drink_tickets, drink_tickets_tbd, status, notes)
  values (ev, 'dc1197e9-caa3-4f69-adfb-c2e01cb79821', 'Champion', 500, 4, 0, true, 'pending', null)
  returning id into sp_pd;
  insert into public.sponsorships (event_id, actor_id, tier, cash_amount, entry_tickets, drink_tickets, drink_tickets_tbd, status, notes)
  values (ev, 'a6421b01-1ca2-41f7-a045-e1845667d7b3', 'Champion', 500, 3, 0, true, 'pending', null)
  returning id into sp_tb;
  insert into public.sponsorships (event_id, actor_id, tier, cash_amount, entry_tickets, drink_tickets, drink_tickets_tbd, status, notes)
  values (ev, 'd3d26350-6c81-4c89-95e9-384efb7cb363', 'Champion', 500, 4, 0, true, 'pending',
          'Dr. Patrick Savery & Dr. Arnaldo Martinez')
  returning id into sp_sa;

  -- Savery & Martinez comps: 3 of 4 named. Patrick + Arnaldo are existing
  -- guest records from DI #2 (matched by id); Alice Zhang is new, no email.
  insert into public.guests (full_name, source, notes)
  values ('Alice Zhang', 'control-center', 'DI #3 sponsor comp (Savery & Martinez)')
  returning id into g_alice;

  insert into public.guest_event_attendance (guest_id, event_id, entry_type, sponsorship_id, quantity, source, ticket_type)
  values ('38308b56-e5d8-454d-8d0e-5a0cfa95a4c2', ev, 'sponsor_comp', sp_sa, 1, 'sponsor', 'Sponsor comp'),
         ('10e5f841-7ac1-46cb-9bcb-8c115a315345', ev, 'sponsor_comp', sp_sa, 1, 'sponsor', 'Sponsor comp'),
         (g_alice,                                 ev, 'sponsor_comp', sp_sa, 1, 'sponsor', 'Sponsor comp');
  get diagnostics n = row_count;
  if n <> 3 then raise exception 'expected 3 comp rows, got %', n; end if;

  -- blockers
  insert into public.tasks (event_id, title, description, status, priority, source) values
    (ev, 'Event start/end time',
         'Confirm with Signal. The event row still carries doors 15:00 from DI #1, which is unconfirmed.', 'blocked', 'high', 'manual'),
    (ev, 'Drink ticket count to lock with Signal',
         'Lock the count (DI #2 locked 330) AND the $ each ticket counts toward the $6,000 bar minimum (DI #2: $12 standard, $13 brunch). Both go in Control Center > Drinks.', 'blocked', 'high', 'manual'),
    (ev, 'Drink tickets per Champion and per $40 ticket',
         'Set on each sponsor card and on the $40 tier in Control Center. Until then the drink ledger reads TBD.', 'blocked', 'high', 'manual'),
    (ev, 'Confirm 3 Champion payments received',
         'Pulse Devices, Theresa Berkman, Dr. Savery & Dr. Martinez - $500 each. Mark paid on the sponsor card.', 'blocked', 'high', 'manual'),
    (ev, 'Name for the $15 portion of the $40 ticket',
         'Currently "Production support". Rename on the $40 tier in Control Center > Tickets.', 'blocked', 'medium', 'manual');

  -- planned costs (the hub FORECAST column - never the P&L)
  insert into public.budget_lines (event_id, scope, direction, ledger, category, planned_amount, amount_tbd, label, notes) values
    (ev, 'event', 'expense', 'dance_infusion', 'Venue',                  6000, false, 'Signal - bar minimum',          'Minimum. Drink tickets count toward it.'),
    (ev, 'event', 'expense', 'dance_infusion', 'Venue',                     0, true,  'Signal - tip guarantee',        'Confirm with Signal.'),
    (ev, 'event', 'expense', 'dance_infusion', 'Meals / Event Supplies',    0, true,  'Food / vendor',                 null),
    (ev, 'event', 'expense', 'dance_infusion', 'Marketing',                 0, true,  'Marketing / print',             null);
end $$;

commit;

-- DOWN (by hand, DI #3 only):
--   delete from tasks where event_id = <ev> and status = 'blocked' and source = 'manual' and title in (...5 above);
--   delete from budget_lines where event_id = <ev> and scope = 'event' and label in (...4 above);
--   delete from guest_event_attendance where event_id = <ev> and entry_type = 'sponsor_comp';
--   delete from guests where full_name = 'Alice Zhang' and source = 'control-center';
--   delete from sponsorships where event_id = <ev>;
--   delete from event_ticket_tiers where event_id = <ev>; delete from event_control where event_id = <ev>;
--   update events set bar_minimum = null where id = <ev>;
