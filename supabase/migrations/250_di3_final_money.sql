-- 250: DI #3 final money (Keith, 2026-11-18).
--
-- What Keith said:
--   * The $6,000 Signal bar minimum was NOT met; the bar rang about $3,000.
--   * Signal keeps the deposit in full. The one $3,000 payment on 11/4 was
--     $2,500 deposit + $500 prepaid drink tickets. The $3,000 "balance" is NOT owed.
--   * $500 tip guarantee (already booked, owner-paid by Venmo) is final.
--   * His owner advance was $1,750 (recorded as $1,500), and it is not coming back:
--     it converts to a founder donation, the DI #1/#2 pattern
--     (third_party_donations, processor 'Founder out-of-pocket', which counts in
--     money raised but NOT in % to the mission).
--   * He will put in $1,300 more for Emmy Adelle ($800) + Frick Frack ($500), also a
--     donation. NOT booked here: it hasn't moved yet. Book it when it does.

begin;

-- guard: the rows we touch are the ones we read on 2026-11-18
do $$ begin
  if (select count(*) from expenses where id = '703e6b56-be2e-4aa7-a000-a471d64071cf'
        and amount = 3000 and status = 'invoiced' and deleted_at is null) <> 1
  then raise exception 'Signal balance row not as expected'; end if;
  if (select count(*) from owner_advances where id = 'e16e6ded-8a08-463d-afb7-7da31b5fd2f6'
        and status = 'contingent' and amount = 1500) <> 1
  then raise exception 'owner advance not as expected'; end if;
  if exists (select 1 from third_party_donations
              where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b'
                and payment_processor = 'Founder out-of-pocket')
  then raise exception 'founder donation already exists on DI #3'; end if;
end $$;

-- 1. The $3,000 balance is not owed (soft delete, reason kept on the row)
update expenses
   set deleted_at = now(),
       description = description || ' | REMOVED 2026-11-18 (Keith): not owed. The $6,000 bar minimum was not met (bar rang ~$3,000); Signal keeps the deposit in full instead of billing the balance.'
 where id = '703e6b56-be2e-4aa7-a000-a471d64071cf';

-- 2. The 11/4 payment: what it actually was
update expenses
   set description = 'Signal - paid 11/4 from DI Bluevine (funded by Keith''s advance + existing DI funds from Zeffy payouts): $2,500 F&B deposit, FORFEITED in full because the $6,000 bar minimum was not met (bar rang ~$3,000), plus $500 prepaid drink tickets. The $3,000 balance on the 09/24/2026 invoice (175 Morgan Ave LLC -> Emergin Leaders Alliance NYC) is not owed. Final per Keith 2026-11-18.'
 where id = '77b92680-cc46-47aa-8c70-d1d42e65778c';

-- 3. Budget lines to their final values (originals kept in notes)
update budget_lines
   set planned_amount = 3000, realized_expense_id = '77b92680-cc46-47aa-8c70-d1d42e65778c',
       realized_at = now(),
       notes = 'FINAL 2026-11-18: minimum was $6,000, NOT met (bar rang ~$3,000). Signal kept the $2,500 deposit + $500 drink tickets ($3,000 paid 11/4); the $3,000 balance is not owed.'
 where id = 'ffc544a7-8de6-494f-9a39-f9ef634007fc';

update budget_lines
   set planned_amount = 500,
       notes = 'FINAL 2026-11-18: budgeted $800; $500 paid by Keith (Venmo) met the guarantee.'
 where id = '8d668563-b937-404b-a759-2be23a82405c';

update budget_lines
   set amount_tbd = false, planned_amount = 0,
       notes = 'FINAL 2026-11-18: no cost recorded for this line.'
 where id in ('8d45ef57-80f0-4019-99f6-03fdf1621097', '098d4da0-8a0f-46aa-a04f-99856132b0c5');

-- 4. Venue terms: record the outcome
update event_control
   set notes = 'Signal NYC, 175 Morgan Ave. Bar minimum $6,000 - NOT MET: bar rang ~$3,000 (Keith, 2026-11-18). Signal kept the $2,500 deposit in full; the $500 drink-ticket prepay (39 tickets) was part of the same $3,000 payment; no balance owed.',
       updated_at = now()
 where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';

-- 5. Owner advance: corrected to $1,750 and converted (a recorded decision, per 225)
update owner_advances
   set amount = 1750, status = 'converted_to_donation', resolved_at = now(),
       resolved_by = 'bd71f88d-fd0a-4e3a-a712-7ac958318c8b',
       resolution_note = 'Keith 2026-11-18: the advance was $1,750 (recorded as $1,500 on 11/5); $500 of it covered the drink tickets inside the $3,000 Signal payment. Minimum not met, so it is not repaid: converted to a founder donation (third_party_donations, Founder out-of-pocket). NB the 11/5 note said the 11/4 transfer was $2,000 incl. the $500 Champion sponsorship; check the DI Bluevine statement for the extra $250.',
       updated_at = now()
 where id = 'e16e6ded-8a08-463d-afb7-7da31b5fd2f6';

insert into third_party_donations (event_id, donor_name, amount, payment_processor, date, notes, actor_id)
values ('f13bf700-7383-497b-91b8-d1f6781e077b', 'Keith Berkman', 1750, 'Founder out-of-pocket', '2026-11-04',
        'Founder donation: Keith''s $1,750 owner advance (owner_advances e16e6ded), converted 2026-11-18 when the Signal minimum was not met. Counts in money raised, not in % to the mission.',
        'ca4a6e21-de18-45d5-a4fd-1d370b6603bb');

-- post-checks
do $$ declare c numeric; d numeric; begin
  select sum(amount) into c from expenses
   where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b' and deleted_at is null;
  if c <> 4979.12 then raise exception 'DI #3 costs %, expected 4979.12', c; end if;
  select sum(amount) into d from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';
  if d <> 2090 then raise exception 'DI #3 donations %, expected 2090', d; end if;
end $$;

commit;
