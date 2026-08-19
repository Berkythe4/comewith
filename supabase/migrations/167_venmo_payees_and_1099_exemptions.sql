-- ============================================================
-- COME WITH — 167 stop treating a payment rail as a payee; clear the easy 1099s
--
-- 1. VENMO IS NOT A PAYEE. Two charges were filed under vendor 'Venmo', which
--    aggregated to $650 and tripped the $600 1099 threshold in v_contractor_1099.
--    They went to two different people for two different things:
--
--        2026-06-22  $400  subwoofer   — buying gear from a private seller
--        2026-07-07  $250  photo shoot — services from a person
--
--    Neither recipient was paid $600, so the aggregate was a false positive
--    caused by grouping on how the money moved rather than who received it. The
--    same mistake in the other direction is worse: if Keith ever Venmos one
--    person repeatedly, their real total hides inside a bucket labelled 'Venmo'.
--
--    Renaming them apart fixes the grouping now. The photo shoot recipient is
--    still unidentified and that is a genuine gap, so it is named as one rather
--    than quietly resolved.
--
-- 2. THE FESTIVAL IS NOT A CONTRACTOR. Elements Music & Arts Festival is $3,349
--    of tickets, passes and on-site costs — buying admission from a festival
--    business, not paying a person for work done for Come With. Exempt, with the
--    reason recorded so the next reviewer does not re-litigate it.
--
-- What deliberately stays 'undecided' after this: 19th & 7th Productions
-- ($1,800) and Janelle Sochet ($900). Both are real service payments over the
-- threshold, and whether a form is owed depends on their entity type, which is a
-- question for them and the accountant — not something to assume here.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Split the Venmo rows by who was actually paid
-- ---------------------------------------------------------------
update public.expenses
   set vendor = 'Private seller (Venmo)',
       description = 'Subwoofer bought secondhand. Goods from an individual — '
                  || 'not 1099 reportable regardless of amount.'
 where deleted_at is null
   and vendor = 'Venmo'
   and date = '2026-06-22'
   and amount = 400.00;

update public.expenses
   set vendor = 'Photographer — name needed (Venmo)',
       description = 'Photo shoot. Services from an individual. Recipient not yet '
                  || 'identified; needed before year end in case of further payments.'
 where deleted_at is null
   and vendor = 'Venmo'
   and date = '2026-07-07'
   and amount = 250.00;

-- ---------------------------------------------------------------
-- 2. Festival admission is not contractor spend
-- ---------------------------------------------------------------
update public.actors
   set tax_1099_status = 'exempt',
       tax_1099_note   = 'Festival business — tickets, passes and on-site costs, '
                      || 'not services performed for Come With'
 where deleted_at is null
   and tax_1099_status is null
   and display_name ilike '%elements%';

commit;

-- DOWN:
--   restore vendor='Venmo' on both rows;
--   set tax_1099_status = null where display_name ilike '%elements%'.
