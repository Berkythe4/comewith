-- =============================================================================
-- 226_di3_costs_and_confirmations.sql - Keith's answers of 2026-11-05 (DI #3)
-- =============================================================================
-- Costs:
--   * Frick Frack BlackJack at DI #3: $500 to Shane Fonda (NEW actor, a person,
--     org = the existing "Frick Frack BlackJack" actor). Not yet paid -> accrued.
--   * Marketing $150 (Meta Ads) already paid by Keith PERSONALLY -> paid,
--     funded_by 'owner' / cash_source 'personal' (the existing pattern for
--     owner-paid business costs). Dated today; the card charges' own dates are
--     not known.
-- Comps / people:
--   * Ethan Pollak takes 1 of Pulse's 4 comps.
--   * "Terri" = Theresa McGuinness, who already has a guest record (DI #1
--     backfill, with an email). The comp moves to that record and the bare
--     "Terri" row 224 created is deleted (nothing else references it).
--     "Bill" = Bill McGuinness (Keith wrote "McGuinnes"; the existing record
--     spells it McGuinness).
--   * Lorena Aragon -> Lorena Hunter (married name). Guest + subscriber name.
--   * Pat Donnelly (the $50 donor) gets a guest record so his RSVP can link.
-- RSVPs: Keith confirmed every suggested match except JAS (stays a suggestion);
--   "Pat" = Pat Donnelly.
-- =============================================================================

begin;

do $$
declare
  ev    uuid;
  sp_pd uuid; sp_tb uuid;
  a_sf  uuid;
  g_pat uuid;
  n int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 not found'; end if;
  if exists (select 1 from public.expenses where external_ref in ('di3:frick-frack-blackjack', 'di3:meta-ads-keith')) then
    raise exception '226 already applied?';
  end if;
  select id into sp_pd from public.sponsorships where event_id = ev and actor_id = 'dc1197e9-caa3-4f69-adfb-c2e01cb79821';
  select id into sp_tb from public.sponsorships where event_id = ev and actor_id = 'a6421b01-1ca2-41f7-a045-e1845667d7b3';

  -- ---- costs
  insert into public.actors (kind, display_name, org_id, notes)
  values ('person', 'Shane Fonda', '1906b1b8-f2bd-490e-b3c9-81177d54ab30', 'Frick Frack BlackJack. Paid $500 for DI #3.')
  returning id into a_sf;

  insert into public.expenses (date, event_id, amount, expected_amount, category, vendor, vendor_actor_id, description,
                               ledger, cash_source, funded_by, status, settled_at, external_ref)
  values
    ('2026-11-14', ev, 500, 500, 'Entertainment / Event', 'Shane Fonda', a_sf,
     'Frick Frack BlackJack at DI #3 - $500 to Shane Fonda. Not yet paid.',
     'dance_infusion', null, 'business', 'accrued', null, 'di3:frick-frack-blackjack'),
    ('2026-11-05', ev, 150, 150, 'Marketing', 'Meta', '92be04e6-7cbb-47ba-aa2e-0f0117970eb0',
     'Meta Ads for DI #3 - paid by Keith personally (owner-funded, not yet reimbursed). Charge dates not recorded.',
     'dance_infusion', 'personal', 'owner', 'paid', now(), 'di3:meta-ads-keith');
  get diagnostics n = row_count;
  if n <> 2 then raise exception 'expected 2 expenses, got %', n; end if;

  -- ---- Ethan Pollak: 1 of Pulse's 4
  insert into public.guest_event_attendance (guest_id, event_id, entry_type, sponsorship_id, quantity, source, ticket_type)
  values ('a9404baa-18a5-4d3e-8bf5-23747a222cde', ev, 'sponsor_comp', sp_pd, 1, 'sponsor', 'Sponsor comp');

  -- ---- Terri -> Theresa McGuinness (existing record), Bill -> Bill McGuinness
  if exists (select 1 from public.guest_event_attendance where event_id = ev and guest_id = 'b23a9e1d-9809-43ff-bae4-dffd83a39d74') then
    raise exception 'Theresa McGuinness already on DI #3';
  end if;
  update public.guest_event_attendance set guest_id = 'b23a9e1d-9809-43ff-bae4-dffd83a39d74'
   where event_id = ev and sponsorship_id = sp_tb and guest_id = '25d9bb97-9860-427f-8003-e6a6f0a75539';
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Terri comp row not found'; end if;
  if exists (select 1 from public.guest_event_attendance where guest_id = '25d9bb97-9860-427f-8003-e6a6f0a75539')
     or exists (select 1 from public.ticketing where guest_id = '25d9bb97-9860-427f-8003-e6a6f0a75539')
     or exists (select 1 from public.subscribers where guest_id = '25d9bb97-9860-427f-8003-e6a6f0a75539')
     or exists (select 1 from public.event_rsvps where '25d9bb97-9860-427f-8003-e6a6f0a75539' in (guest_id, suggested_guest_id))
  then raise exception 'the 224 "Terri" row is still referenced'; end if;
  delete from public.guests where id = '25d9bb97-9860-427f-8003-e6a6f0a75539' and source = 'control-center:224' and full_name = 'Terri';

  update public.guests set full_name = 'Bill McGuinness',
         notes = 'Theresa Berkman comp (DI #3). With Theresa McGuinness. Surname confirmed by Keith 2026-11-05.'
   where id = 'c154314c-b6bb-415d-93a5-8762bdeccb51' and full_name = 'Bill';
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Bill row not found'; end if;

  -- ---- Lorena Hunter (married name)
  update public.guests set full_name = 'Lorena Hunter',
         notes = concat_ws(' ', notes, 'Formerly Lorena Aragon (married name; Keith 2026-11-05).')
   where id = '8421da7a-d2c4-408b-8553-ff14e6dff7f8' and full_name = 'Lorena Aragon';
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Lorena guest not found'; end if;
  update public.subscribers set full_name = 'Lorena Hunter' where id = '09da6f05-2d00-43a4-93c1-b6161b47d778' and full_name = 'Lorena Aragon';

  -- ---- Pat Donnelly (donor) gets a record so the RSVP can link
  insert into public.guests (full_name, email, source, notes)
  values ('Pat Donnelly', 'pmd55@optonline.net', 'control-center:226', 'DI #3 donor ($50, Zeffy 10/23). Partiful "Pat".')
  returning id into g_pat;

  -- ---- RSVPs: confirmed suggestions become links (JAS stays a suggestion)
  update public.event_rsvps
     set guest_id = suggested_guest_id, suggested_guest_id = null,
         match_note = 'Confirmed by Keith 2026-11-05 (was: ' || match_note || ')'
   where event_id = ev and suggested_guest_id is not null
     and external_key in ('partiful:arthemise', 'partiful:lovely', 'partiful:jessica', 'partiful:irene',
                          'partiful:kierney', 'partiful:martin kearney', 'partiful:scott (rainbow tutu)');
  get diagnostics n = row_count;
  if n <> 7 then raise exception 'expected 7 confirmed RSVP links, got %', n; end if;
  update public.event_rsvps set guest_id = g_pat, match_note = 'Confirmed by Keith 2026-11-05: Pat Donnelly (the $50 donor)'
   where event_id = ev and external_key = 'partiful:pat' and guest_id is null;
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Pat RSVP not found'; end if;

  if (select count(*) from public.event_rsvps where event_id = ev and guest_id is not null) <> 16 then raise exception 'expected 16 linked RSVPs'; end if;
  if (select count(*) from public.guest_event_attendance where sponsorship_id = sp_pd) <> 1 then raise exception 'Pulse comps'; end if;
end $$;

commit;
