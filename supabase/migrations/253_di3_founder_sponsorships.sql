-- 253: DI #3 founder money is SPONSORSHIP, not donation (Keith, 2026-11-18).
--
-- Keith: "Can't call my $1,750 or Come With $150 a donation because that implies it
-- goes to MS Society, must call it sponsorship ... keep it separate from the original
-- 4 sponsors." On a fundraiser a DONATION reads as money for the beneficiary; this
-- money covered production costs. So:
--   * the three 'Founder out-of-pocket' third_party_donations rows from 250-252 are
--     removed (Keith $1,250 advance, Keith $500 tip, Come With $150 Meta ad);
--   * they become two sponsorships with tier 'Founder sponsorship' - a tier none of
--     the original four (Pulse, Theresa, Savery/Martinez, Keith's Champion) carries;
--   * owner_advances gains the status 'converted_to_sponsorship' (CHECK widened,
--     additive) and the advance moves to it.
-- Money raised and % to the mission are unchanged: neither founder donations nor
-- sponsor cash count as mission money (comp_mission_value is null on DI #3).
-- Come With had no actor row (searched 'come with' / 'comewith'); one org actor is
-- created so the sponsorship has a name in the Control Center.

begin;

do $$ begin
  if (select count(*) from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
        and payment_processor ilike 'Founder out-of-pocket%') <> 3
  then raise exception 'expected 3 founder donation rows'; end if;
  if (select sum(amount) from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
        and payment_processor ilike 'Founder out-of-pocket%') <> 1900
  then raise exception 'founder donations do not total 1900'; end if;
  if exists (select 1 from actors where display_name ilike 'come with' and deleted_at is null)
  then raise exception 'a Come With actor already exists'; end if;
end $$;

alter table public.owner_advances drop constraint owner_advances_status_check;
alter table public.owner_advances add constraint owner_advances_status_check
  check (status = any (array['contingent','owed','repaid','converted_to_donation',
                             'converted_to_sponsorship','written_off']));

update owner_advances
   set status = 'converted_to_sponsorship', updated_at = now(),
       resolution_note = 'Keith 2026-11-18: minimum not met, not repaid; converted to Keith''s FOUNDER SPONSORSHIP on DI #3 (not a donation - a donation reads as money for the MS Society). Amount per the DI Bluevine export: one transfer in on 11/4, "Wealthfront, EDI PYMNTS" $1,750 = $500 Champion sponsorship (the drinks) + $1,250 advance. Recorded $1,500 on 11/5, $1,750 in 250 (double count), corrected in 251.'
 where id = 'e16e6ded-8a08-463d-afb7-7da31b5fd2f6';

delete from third_party_donations
 where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
   and payment_processor ilike 'Founder out-of-pocket%';

insert into actors (kind, display_name, notes)
values ('org', 'Come With', 'Come With itself (the business), as a counterparty - e.g. its in-kind founder sponsorship of DI #3 (253).');

insert into sponsorships (event_id, actor_id, tier, cash_amount, status, notes, payment_refs)
values
  ('f13bf700-7383-497b-91b8-d1f6781e077b', 'ca4a6e21-de18-45d5-a4fd-1d370b6603bb', 'Founder sponsorship', 1750, 'paid',
   'Keith''s founder sponsorship of DI #3, separate from his $500 Champion package. $1,250 = his 11/4 owner advance (owner_advances e16e6ded), not repaid when the Signal minimum was missed; $500 = the Signal tip guarantee he paid by Venmo on 11/14 (expense e0e94bc1). Covers production costs; NOT a donation and not money to the MS Society. Keith 2026-11-18.',
   array['owner_advance:e16e6ded-8a08-463d-afb7-7da31b5fd2f6', 'expense:e0e94bc1-0e9d-41a1-b00c-7525d521e4f5']),
  ('f13bf700-7383-497b-91b8-d1f6781e077b', (select id from actors where display_name = 'Come With' and kind = 'org'), 'Founder sponsorship', 150, 'paid',
   'Come With''s in-kind founder sponsorship of DI #3: it paid the $150 DI #3 Meta ad from its own Bluevine on 10/16 (expense 43c972fa, on Come With''s books). NOT a donation and not money to the MS Society. Keith 2026-11-18.',
   array['expense:43c972fa-5dc7-45fa-b5ac-b32421ef80a4']);

update expenses
   set description = replace(description, 'Keith donated it (founder donation, third_party_donations).', 'it is part of Keith''s $1,750 founder sponsorship of DI #3 (sponsorships, 253).')
 where id = 'e0e94bc1-0e9d-41a1-b00c-7525d521e4f5';

update expenses
   set description = 'Meta Ads for DI #3, paid from Come With''s Bluevine (stays on Come With''s books). Come With''s in-kind founder sponsorship of DI #3 (sponsorships, 253). Keith 2026-11-18.'
 where id = '43c972fa-5dc7-45fa-b5ac-b32421ef80a4';

do $$ declare d numeric; s numeric; n int; begin
  select coalesce(sum(amount),0) into d from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';
  select sum(cash_amount), count(*) into s, n from sponsorships
   where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b' and status = 'paid';
  if d <> 340 or s <> 3900 or n <> 6 then
    raise exception 'donations % sponsors % (% rows), expected 340 / 3900 / 6', d, s, n; end if;
end $$;

commit;
