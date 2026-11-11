-- =============================================================================
-- 241_emmy_adelle_leif_kseniya_janelle.sql - Keith, 2026-11-11 (DI #3)
-- =============================================================================
--   * "Emmy Adelle" is the correct spelling (handle @emmyyadelle). Renames the
--     actor and the guest record, plus the DI #3 door-list text that names her.
--     Social posts are NOT touched: all four that say "Emmy Adele" are already
--     posted, and two of them record that the DI3 flier misspelled her.
--   * Leif Fosse gets an actor record: Emmy Adelle's manager. actor_roles has no
--     'manager' role, so the relationship lives in notes.
--   * Kseniya Sovenko gets an actor record, linked (org_id) to a NEW org actor
--     "NYC Rave Girls", which she co-founded (one of two founders).
--   * Janelle Sochet (Marketing Director, sub_admin) goes on the DI #3 door list
--     as staff, on her existing guest record (already linked to her actor),
--     and on the crew like Henry (228). Not a comp; drink tickets TBD.
-- =============================================================================

begin;

do $$
declare
  ev      uuid;
  a_leif  uuid; a_ks uuid; a_nrg uuid;
  g_leif  constant uuid := 'e1a0ca21-b563-4e64-a5d8-5ef261597413';
  g_ks    uuid;
  n int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 not found'; end if;

  -- ---- Emmy Adelle
  update public.actors set display_name = 'Emmy Adelle',
         notes = replace(coalesce(notes, ''), 'Emmy Adele', 'Emmy Adelle')
   where id = 'b7abac4e-9409-4812-b0f3-eb1481e5190c' and display_name = 'Emmy Adele';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Emmy actor (already renamed?)'; end if;
  update public.guests set full_name = 'Emmy Adelle'
   where id = '76d7d8e6-9394-446b-8e47-a989d1f41139' and full_name = 'Emmy Adele';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Emmy guest'; end if;
  update public.guest_event_attendance
     set notes = replace(notes, 'Emmy Adele', 'Emmy Adelle'),
         ticket_type = replace(ticket_type, 'Emmy Adele', 'Emmy Adelle')
   where event_id = ev and (notes like '%Emmy Adele%' or ticket_type like '%Emmy Adele%');
  get diagnostics n = row_count; if n <> 3 then raise exception 'expected 3 door-list rows naming Emmy, got %', n; end if;
  update public.guests set notes = replace(notes, 'Emmy Adele', 'Emmy Adelle')
   where source = 'control-center:240' and notes like '%Emmy Adele%';
  get diagnostics n = row_count; if n <> 2 then raise exception 'expected Leif + Abby guest notes, got %', n; end if;

  -- ---- Leif Fosse: Emmy's manager
  insert into public.actors (kind, display_name, notes)
  values ('person', 'Leif Fosse', 'Manager of Emmy Adelle. Comp of Emmy at DI #3 (Keith 2026-11-11).')
  returning id into a_leif;
  update public.guests set actor_id = a_leif where id = g_leif and full_name = 'Leif Fosse' and actor_id is null;
  get diagnostics n = row_count; if n <> 1 then raise exception 'Leif guest'; end if;

  -- ---- Kseniya Sovenko: co-founder of NYC Rave Girls
  if exists (select 1 from public.actors where display_name ilike 'nyc rave girls' and deleted_at is null) then
    raise exception 'NYC Rave Girls actor already exists';
  end if;
  insert into public.actors (kind, display_name, notes)
  values ('org', 'NYC Rave Girls', 'NYC rave collective; two founders, one of them Kseniya Sovenko (Keith 2026-11-11).')
  returning id into a_nrg;
  insert into public.actors (kind, display_name, org_id, notes)
  values ('person', 'Kseniya Sovenko', a_nrg, 'One of the two founders of NYC Rave Girls. Keith comp (2 tickets) at DI #3.')
  returning id into a_ks;
  update public.guests set actor_id = a_ks
   where source = 'control-center:240' and full_name = 'Kseniya Sovenko' and actor_id is null
  returning id into g_ks;
  if g_ks is null then raise exception 'Kseniya guest'; end if;

  -- ---- Janelle Sochet: staff on the door list + crew
  if exists (select 1 from public.guest_event_attendance where event_id = ev and guest_id = 'a34bd1e2-a6e8-4778-b3a1-d7afba944087') then
    raise exception 'Janelle already on the DI #3 door list';
  end if;
  insert into public.guest_event_attendance (guest_id, event_id, entry_type, quantity, source, ticket_type, notes)
  values ('a34bd1e2-a6e8-4778-b3a1-d7afba944087', ev, 'staff', 1, 'manual', 'Crew', 'Janelle Sochet - Marketing Director. Drink tickets TBD.');
  if not exists (select 1 from public.event_participants where event_id = ev and actor_id = '737162ab-bed3-47b8-b0f0-a43fb204b9ad') then
    insert into public.event_participants (event_id, actor_id, role, roles, notes)
    values (ev, '737162ab-bed3-47b8-b0f0-a43fb204b9ad', 'crew', '{crew}', 'Marketing Director.');
  end if;

  -- ---- verify
  if exists (select 1 from public.guest_event_attendance where event_id = ev and (notes like '%Emmy Adele %' or notes like '%Emmy Adele.%' or ticket_type like '%Emmy Adele)%'))
    then raise exception 'an "Emmy Adele" survived on the door list'; end if;
  if (select coalesce(sum(greatest(1, quantity)), 0) from public.guest_event_attendance
       where event_id = ev and entry_type in ('dj', 'artist', 'comp', 'sponsor_comp')) <> 20 then raise exception 'comp count moved off 20'; end if;
  if (select count(*) from public.guest_event_attendance where event_id = ev and entry_type = 'staff') <> 2 then raise exception 'staff <> 2 (Henry + Janelle)'; end if;
  if (select count(*) from (select guest_id from public.guest_event_attendance where event_id = ev
       and entry_type <> 'ticket' group by guest_id having count(*) > 1) d) <> 0 then raise exception 'someone is on the door list twice'; end if;
end $$;

commit;
