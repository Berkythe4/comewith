-- =============================================================================
-- 240_di3_final_comps_and_drinks.sql - Keith, 2026-11-11 (DI #3, FINAL comp list)
-- =============================================================================
-- Replaces every prior DI #3 comp entry with Keith's final list.
--
-- 1. Off the comps:
--    * Theresa McGuinness ("Terry", tess8897@aol.com) and Bill McGuinness come off
--      Theresa Berkman's comps. Terry is a paid Zeffy buyer (2 x $40). Her Zeffy
--      order is NOT in prod yet, so it arrives by re-running the Zeffy import
--      (dedupes on receipt #). It is not hand-entered here because a hand row
--      would be booked a second time on that import.
--    * Theresa Berkman Champion: comps 3 -> 1 (Theresa herself).
--    * Keith Berkman Champion: comps 3 -> 0 (Keith is on the list as a DJ).
-- 2. Comps exactly as listed:
--    * DJs: the 5 existing dj rows, 1 each, drinks TBD (unchanged).
--    * Performers (Keith's): JayJay (2 -> 1) + Amanda Maldonado move off Keith's
--      sponsorship to entry_type 'artist' (lineup). Amanda's +1s pay.
--    * Artist guests (entry_type 'comp', source in ticket_type/notes):
--      Leif Fosse + Abby Cavallo (Emmy Adele), Kierney Kearney (KRNeY).
--      Kierney = the existing "Kierney Brennan" record (krosie6290@), which
--      Keith confirmed as Partiful "Kierney" in 226. Renamed to her married name,
--      with the old name kept in notes (same pattern as Lorena Hunter, 226).
--    * Keith comps: Amanda Brundige (existing record b1a6e0dd, already spelled
--      Brundige and linked to her Partiful RSVP; there is no "Brunidge" record in
--      prod), Angela Tabone (existing record; her $50 Zeffy donation stays
--      separate and is not in prod yet either), Sonia Sukdeo (NEW guest, linked
--      to the DJ Sauci Soni actor who played DI #2), Kseniya Sovenko x2 (NEW).
--    * Sponsor comps: Theresa Berkman 1; Savery/Martinez 4 (Patrick, Arnaldo,
--      Alice, 1 open, unchanged); Pulse 4 = Ethan + 3 UNCONFIRMED (shown as open
--      slots, which never count as heads).
-- 3. Drink tickets: 38 locked with Signal. Packages + handed-out on the guest
--    row they sit on (the ledger counts max(package, handed), so never twice):
--    Theresa 6 (Theresa Berkman), Pulse 10 (Ethan Pollak), Savery/Martinez 4
--    (Patrick Savery). Everyone else TBD. Allocated 20 / 38 / 18 unallocated.
-- =============================================================================

begin;

do $$
declare
  ev    uuid;
  sp_tb uuid; sp_kb uuid; sp_pd uuid; sp_ps uuid;
  g_leif uuid; g_abby uuid; g_sonia uuid; g_ks uuid;
  g_kier constant uuid := 'a42859cb-64e9-4719-a86b-360c66ea7b26';
  n int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 not found'; end if;
  select s.id into sp_tb from public.sponsorships s where s.event_id = ev and s.actor_id = 'a6421b01-1ca2-41f7-a045-e1845667d7b3';
  select s.id into sp_pd from public.sponsorships s where s.event_id = ev and s.actor_id = 'dc1197e9-caa3-4f69-adfb-c2e01cb79821';
  select s.id into sp_kb from public.sponsorships s join public.actors a on a.id = s.actor_id
   where s.event_id = ev and a.display_name = 'Keith Berkman';
  select s.id into sp_ps from public.sponsorships s join public.actors a on a.id = s.actor_id
   where s.event_id = ev and a.display_name = 'Patrick Savery';
  if sp_tb is null or sp_pd is null or sp_kb is null or sp_ps is null then raise exception 'a sponsorship is missing'; end if;
  if exists (select 1 from public.guests where source = 'control-center:240') then raise exception '240 already applied?'; end if;

  -- ---- 1. Terry + Bill off Theresa Berkman's comps
  delete from public.guest_event_attendance
   where event_id = ev and sponsorship_id = sp_tb
     and guest_id in ('b23a9e1d-9809-43ff-bae4-dffd83a39d74', 'c154314c-b6bb-415d-93a5-8762bdeccb51');
  get diagnostics n = row_count;
  if n <> 2 then raise exception 'expected Terry + Bill comp rows, got %', n; end if;
  update public.guests set notes = concat_ws(' ', notes, 'Not a DI #3 comp (Keith 2026-11-11); paid buyer on Zeffy, 2 x $40.')
   where id = 'b23a9e1d-9809-43ff-bae4-dffd83a39d74';
  update public.guests set notes = concat_ws(' ', notes, 'Removed from DI #3 comps 2026-11-11.')
   where id = 'c154314c-b6bb-415d-93a5-8762bdeccb51';

  update public.sponsorships
     set entry_tickets = 1, drink_tickets = 6, drink_tickets_tbd = false,
         notes = 'Paid via Zeffy 10/7/2026 ($500). 1 comp: Theresa Berkman (final list 2026-11-11; Terry + Bill McGuinness removed - Terry bought 2 tickets on Zeffy). 6 drink tickets, on her guest record.'
   where id = sp_tb;

  -- ---- Keith's Champion: 0 comps; JayJay + Amanda become performers
  update public.sponsorships
     set entry_tickets = 0,
         notes = 'Paid 11/4/2026 via Bluevine (DI): $500 of Keith''s $2,000 personal transfer. Not Zeffy. 0 comps - Keith is on the list as a DJ; his personal comps (performers, guests) are plain comps sourced to Keith (final list 2026-11-11).'
   where id = sp_kb;
  update public.guest_event_attendance
     set entry_type = 'artist', sponsorship_id = null, source = 'manual', quantity = 1,
         ticket_type = 'Performer (Keith)', notes = 'Flow performer. Comp from Keith - 1 entry (final list 2026-11-11).'
   where event_id = ev and sponsorship_id = sp_kb and guest_id = 'e741ca25-301f-4f54-acf7-17e5a2928628';
  get diagnostics n = row_count; if n <> 1 then raise exception 'JayJay row'; end if;
  update public.guest_event_attendance
     set entry_type = 'artist', sponsorship_id = null, source = 'manual', quantity = 1,
         ticket_type = 'Performer (Keith)', notes = 'Flow performer. Comp from Keith - 1 entry. Her +1s pay (not comped). Final list 2026-11-11.'
   where event_id = ev and sponsorship_id = sp_kb and guest_id = 'f732b371-c436-4f1a-a222-195573c1d7bb';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Amanda Maldonado row'; end if;

  -- ---- Pulse: Ethan confirmed, 3 tentative; 10 drinks on Ethan
  update public.sponsorships
     set drink_tickets = 10, drink_tickets_tbd = false,
         notes = 'Paid via Zeffy 9/17/2026 ($500, card, Ethan Pollak). 4 comps: Ethan Pollak confirmed; up to 3 guests UNCONFIRMED (open slots, not counted). 10 drink tickets on Ethan''s record until Pulse guests are confirmed, then redistribute.'
   where id = sp_pd;
  update public.guest_event_attendance set drink_tickets = 10
   where event_id = ev and sponsorship_id = sp_pd and guest_id = 'a9404baa-18a5-4d3e-8bf5-23747a222cde';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Ethan row'; end if;

  -- ---- Savery/Martinez: 4 drinks on Patrick
  update public.sponsorships
     set drink_tickets = 4, drink_tickets_tbd = false,
         notes = notes || ' 4 drink tickets on Patrick Savery''s record until Keith says how to split (2026-11-11).'
   where id = sp_ps;
  update public.guest_event_attendance set drink_tickets = 4
   where event_id = ev and sponsorship_id = sp_ps and guest_id = '38308b56-e5d8-454d-8d0e-5a0cfa95a4c2';
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Patrick Savery row'; end if;

  -- ---- Theresa Berkman: 6 drinks on her own row
  update public.guest_event_attendance set drink_tickets = 6
   where event_id = ev and sponsorship_id = sp_tb and guest_id = '7b8ab30c-16e9-4411-9404-fe8acecd8e03';
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Theresa Berkman row'; end if;

  -- ---- 2. Artist guests
  insert into public.guests (full_name, source, notes)
  values ('Leif Fosse', 'control-center:240', 'Comp of Emmy Adele (DI #3).') returning id into g_leif;
  insert into public.guests (full_name, source, notes)
  values ('Abby Cavallo', 'control-center:240', 'Comp of Emmy Adele (DI #3).') returning id into g_abby;

  update public.guests set full_name = 'Kierney Kearney',
         notes = concat_ws(' ', notes, 'Formerly Kierney Brennan (married name). KRNeY''s wife. Keith 2026-11-11.')
   where id = g_kier and full_name = 'Kierney Brennan';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Kierney guest'; end if;
  update public.subscribers set full_name = 'Kierney Kearney'
   where id = 'f4fce6c0-d8c7-407b-a983-3fc774ce0fdb' and full_name = 'Kierney Brennan';

  insert into public.guest_event_attendance (guest_id, event_id, entry_type, quantity, source, ticket_type, notes) values
    (g_leif, ev, 'comp', 1, 'manual', 'Comp (Emmy Adele)', 'Artist guest - comp of Emmy Adele. Drink tickets TBD.'),
    (g_abby, ev, 'comp', 1, 'manual', 'Comp (Emmy Adele)', 'Artist guest - comp of Emmy Adele. Drink tickets TBD.'),
    (g_kier, ev, 'comp', 1, 'manual', 'Comp (KRNeY)',      'Artist guest - comp of KRNeY (his wife). Partiful "Kierney". Drink tickets TBD.');

  -- ---- Keith comps
  insert into public.guests (full_name, actor_id, source, notes)
  values ('Sonia Sukdeo', '1c87b083-a56d-4779-89af-b67a271e071c', 'control-center:240',
          'DJ name Sauci Soni (actor "DJ Sauci Soni"). Played DI #2; guest at DI #3, not on the lineup.')
  returning id into g_sonia;
  insert into public.guests (full_name, source, notes)
  values ('Kseniya Sovenko', 'control-center:240', 'Keith comp, DI #3 (2 tickets).') returning id into g_ks;

  insert into public.guest_event_attendance (guest_id, event_id, entry_type, quantity, source, ticket_type, notes) values
    ('b1a6e0dd-a498-445d-8ea9-fb0a225dbe20', ev, 'comp', 1, 'manual', 'Comp (Keith)', 'Keith comp. Spelling is Brundige (not "Brunidge"). Drink tickets TBD.'),
    ('3058ca26-99db-45d4-b8cf-af2ec49eeb00', ev, 'comp', 1, 'manual', 'Comp (Keith)', 'Keith comp. Also donated $50 on Zeffy - the donation is separate. Drink tickets TBD.'),
    (g_sonia,                                ev, 'comp', 1, 'manual', 'Comp (Keith)', 'Keith comp. Sauci Soni - played DI #2, guest at DI #3 (not on the lineup). Drink tickets TBD.'),
    (g_ks,                                   ev, 'comp', 2, 'manual', 'Comp (Keith)', 'Keith comp - 2 tickets. Drink tickets TBD.');

  -- ---- 3. Locked drink count
  update public.event_control set drink_tickets_locked = 38 where event_id = ev;
  get diagnostics n = row_count; if n <> 1 then raise exception 'event_control row'; end if;

  -- ---- 4. Verify
  if (select coalesce(sum(greatest(1, quantity)), 0) from public.guest_event_attendance
       where event_id = ev and entry_type in ('dj', 'artist', 'comp', 'sponsor_comp')) <> 20
  then raise exception 'confirmed comp entries <> 20: %',
    (select sum(greatest(1, quantity)) from public.guest_event_attendance where event_id = ev and entry_type in ('dj', 'artist', 'comp', 'sponsor_comp'));
  end if;
  if (select count(*) from public.guest_event_attendance where event_id = ev and entry_type = 'sponsor_comp') <> 5 then raise exception 'sponsor comps <> 5'; end if;
  if (select count(*) from (select guest_id from public.guest_event_attendance where event_id = ev
       and entry_type <> 'ticket' group by guest_id having count(*) > 1) d) <> 0 then raise exception 'someone is on the door list twice'; end if;
  if (select coalesce(sum(drink_tickets), 0) from public.guest_event_attendance where event_id = ev) <> 20 then raise exception 'handed drinks <> 20'; end if;
  if (select sum(drink_tickets) from public.sponsorships where event_id = ev and not drink_tickets_tbd) <> 20 then raise exception 'drink packages <> 20'; end if;
  if (select entry_tickets from public.sponsorships where id = sp_pd) - 1 <> 3 then raise exception 'Pulse open slots <> 3'; end if;
  if (select entry_tickets from public.sponsorships where id = sp_ps) - (select count(*) from public.guest_event_attendance where sponsorship_id = sp_ps) <> 1
    then raise exception 'Savery open slots <> 1'; end if;
end $$;

commit;
