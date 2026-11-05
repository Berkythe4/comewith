-- =============================================================================
-- 228_di3_henry_crew_keith_comps.sql - Keith, 2026-11-05 (DI #3)
-- =============================================================================
--   * Henry Zaradich joins the crew as PRODUCER: an event_participants row on his
--     existing actor ("Henry") and a staff row on the door list for his existing
--     guest record ("Henry J. Zaradich", already linked to that actor).
--   * JayJay (2 entries) and Amanda Maldonado (1) become KEITH's comps: they move
--     from plain comps onto Keith's Champion sponsorship, whose comp count goes
--     0 -> 3. Comps count ENTRIES, not rows (JayJay is a party of 2) - the
--     Control Center counts the same way from this release.
-- =============================================================================

begin;

do $$
declare
  ev    uuid;
  sp_kb uuid;
  n int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 not found'; end if;
  select s.id into sp_kb from public.sponsorships s join public.actors a on a.id = s.actor_id
   where s.event_id = ev and a.display_name = 'Keith Berkman' and s.entry_tickets = 0;
  if sp_kb is null then raise exception 'Keith Champion not found with 0 comps - already applied?'; end if;

  -- Henry: crew / producer
  if exists (select 1 from public.event_participants where event_id = ev and actor_id = 'e97045d6-aa0b-451c-8711-bc7ef1dfa778') then
    raise exception 'Henry already on the DI #3 lineup';
  end if;
  insert into public.event_participants (event_id, actor_id, role, roles, notes)
  values (ev, 'e97045d6-aa0b-451c-8711-bc7ef1dfa778', 'crew', '{crew,producer}', 'Producer.');
  insert into public.guest_event_attendance (guest_id, event_id, entry_type, quantity, source, ticket_type, notes)
  values ('7e3cc8c1-5d69-440f-8a7e-b13d4542150d', ev, 'staff', 1, 'manual', 'Crew', 'Henry Zaradich - producer.');

  -- JayJay + Amanda -> Keith's comps
  update public.sponsorships
     set entry_tickets = 3,
         notes = 'Paid 11/4/2026 via Bluevine (DI): $500 of Keith''s $2,000 personal transfer. Not Zeffy. 3 comps: JayJay (x2) and Amanda Maldonado. Keith himself is on the list as a DJ.'
   where id = sp_kb;
  update public.guest_event_attendance
     set entry_type = 'sponsor_comp', sponsorship_id = sp_kb, source = 'sponsor', ticket_type = 'Sponsor comp'
   where event_id = ev and entry_type = 'comp'
     and guest_id in (select id from public.guests where source = 'control-center:224' and full_name in ('JayJay', 'Amanda Maldonado'));
  get diagnostics n = row_count;
  if n <> 2 then raise exception 'expected JayJay + Amanda, got % rows', n; end if;
  if (select sum(quantity) from public.guest_event_attendance where sponsorship_id = sp_kb) <> 3 then raise exception 'Keith comps <> 3 entries'; end if;
end $$;

commit;
