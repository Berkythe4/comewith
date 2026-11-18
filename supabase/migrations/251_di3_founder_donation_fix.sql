-- 251: DI #3 founder money, corrected against the DI Bluevine export (Keith, 2026-11-18).
--
-- Evidence: DI Bluevine transactions export (downloaded 2026-11-18) shows ONE transfer
-- in from Keith on 11/4: "Wealthfront, EDI PYMNTS" +$1,750.00. Not $2,000.
-- That $1,750 is ALL of Keith's 11/4 money, and it was already split two ways in the
-- books: his $500 Champion sponsorship (sponsorships, payment_ref
-- bluevine:2026-11-04:keith-transfer) + the advance. 250 booked the advance as $1,750,
-- which counted $500 of the transfer twice. Keith: "$500 of it went to drinks" = the
-- Champion package; "the other $1,250" = the advance. So: advance + donation = $1,250.
--
-- Also (Keith, 2026-11-18):
--   * the $500 Signal tip he paid by Venmo is donated, not repaid -> founder donation;
--   * Tst*signal $29.12 on 11/7 (personal card) was PERSONAL, not a DI cost.
-- The $150 Meta row is NOT touched here: it may be the same ad as the 10/16
-- "FACEBK *AMBJF7ERA4" $150 on Come With's Bluevine. Pending Keith.

begin;

do $$ begin
  if (select amount from owner_advances where id = 'e16e6ded-8a08-463d-afb7-7da31b5fd2f6') <> 1750
  then raise exception 'advance not at 1750'; end if;
  if (select count(*) from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
        and payment_processor = 'Founder out-of-pocket' and amount = 1750) <> 1
  then raise exception 'founder donation row not as expected'; end if;
  if (select count(*) from expenses where id = '0667da96-2149-46a6-9399-a9f60877bb7b'
        and amount = 29.12 and deleted_at is null) <> 1
  then raise exception 'Tst*signal row not as expected'; end if;
  if (select count(*) from expenses where id = 'e0e94bc1-0e9d-41a1-b00c-7525d521e4f5'
        and amount = 500 and reimbursed_at is null) <> 1
  then raise exception 'tip row not as expected'; end if;
end $$;

-- 1. advance: $1,250, not $1,750
update owner_advances
   set amount = 1250, updated_at = now(),
       resolution_note = 'Keith 2026-11-18: converted to a founder donation (minimum not met, not repaid). Amount per the DI Bluevine export: one transfer in on 11/4, "Wealthfront, EDI PYMNTS" $1,750 = $500 Champion sponsorship (the drinks) + $1,250 advance. Was recorded $1,500 on 11/5 (from an assumed $2,000 transfer), then $1,750 in 250 (double-counted the $500); corrected in 251.'
 where id = 'e16e6ded-8a08-463d-afb7-7da31b5fd2f6';

update third_party_donations
   set amount = 1250,
       notes = 'Founder donation: Keith''s $1,250 owner advance (owner_advances e16e6ded), converted 2026-11-18 when the Signal minimum was not met. Part of the $1,750 Wealthfront transfer of 11/4; the other $500 is his Champion sponsorship. Counts in money raised, not in % to the mission.'
 where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
   and payment_processor = 'Founder out-of-pocket' and amount = 1750;

update sponsorships
   set notes = replace(notes, '$500 of Keith''s $2,000 personal transfer', '$500 of Keith''s $1,750 personal transfer (Wealthfront, per the DI Bluevine export; the other $1,250 is the advance)')
 where id = (select s.id from sponsorships s
              where s.event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
                and s.actor_id = 'ca4a6e21-de18-45d5-a4fd-1d370b6603bb');

-- 2. the $500 tip: donated, not repaid
insert into third_party_donations (event_id, donor_name, amount, payment_processor, date, notes, actor_id)
values ('f13bf700-7383-497b-91b8-d1f6781e077b', 'Keith Berkman', 500, 'Founder out-of-pocket', '2026-11-14',
        'Founder donation: the $500 Signal tip guarantee Keith paid personally by Venmo (expense e0e94bc1), donated 2026-11-18 rather than repaid. Counts in money raised, not in % to the mission.',
        'ca4a6e21-de18-45d5-a4fd-1d370b6603bb');

update expenses
   set reimbursed_at = now(),
       description = description || ' | SETTLED 2026-11-18: not repaid in cash - Keith donated it (founder donation, third_party_donations).'
 where id = 'e0e94bc1-0e9d-41a1-b00c-7525d521e4f5';

-- 3. Tst*signal $29.12: personal. Soft-delete and KEEP external_ref - the unique index
--    covers deleted rows, so the card import cannot re-create it.
update expenses
   set deleted_at = now(),
       description = 'REMOVED 2026-11-18 (Keith): personal, not a DI #3 or business cost. Kept (deleted) so the card import does not re-create it.'
 where id = '0667da96-2149-46a6-9399-a9f60877bb7b';

do $$ declare c numeric; d numeric; f numeric; begin
  select sum(amount) into c from expenses
   where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b' and deleted_at is null;
  select sum(amount), sum(amount) filter (where payment_processor = 'Founder out-of-pocket') into d, f
    from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';
  if c <> 4950 then raise exception 'costs %, expected 4950', c; end if;
  if d <> 2090 or f <> 1750 then raise exception 'donations % founder %, expected 2090 / 1750', d, f; end if;
end $$;

commit;
