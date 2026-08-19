-- ============================================================
-- COME WITH — 169 one videographer payment, not two; gear filed as gear
--
-- 1. 19TH & 7TH WAS PAID ONCE, NOT TWICE. Two $900 charges existed:
--
--      2026-01-17  hand-entered 29 May, category Contractors, linked to
--                  Crossroads Café Artist Showcase, described as the
--                  videographer fee
--      2026-02-04  created by today's statement import, vendor
--                  'In *19th & 7th, Inc.', category Production, no event,
--                  no description
--
--    17 days apart is a card settling, not a second engagement. The import's
--    adopt-or-insert check keys on date and amount, and the dates differ, so it
--    inserted rather than adopted.
--
--    The hand-entered row is kept because it carries the event link, the
--    category and the description. But the import row holds the external_ref,
--    which is the only thing stopping the next import inserting it all over
--    again — so the ref is MOVED onto the surviving row first. Delete without
--    moving it and this duplicate comes back on the next Bluevine run.
--
--    The descriptor also settles a 1099 question: 'Inc.' means incorporated,
--    and payments to corporations are not 1099-NEC reportable. Recorded on the
--    actor with the evidence, so it does not get re-asked every January.
--
-- 2. GEAR THAT WAS NOT IN 'EQUIPMENT'. Equipment totalled $10,651.77 while the
--    stolen rig alone cost $11,837.51 — and not everything was stolen, so the
--    category was provably wrong, not merely arguable. Four purchases were
--    sitting elsewhere:
--
--      $1,513.35  Best Buy work laptop        (was Production)
--      $  400.00  subwoofer, private seller   (was Production)
--      $   88.74  external hard drive         (was Operations)
--      $   28.99  Sweetwater cables           (was Supplies)
--
--    Camping gear bought for a festival is deliberately NOT moved — it is
--    travel kit for a networking trip, not production equipment, and folding it
--    in would overstate the depreciable asset base.
--
--    Total spend does not change. Only the category does, which is what the
--    equipment and theft schedule is read from.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Carry the import identity across, then drop the duplicate
-- ---------------------------------------------------------------
-- external_ref carries a unique index, so it cannot sit on both rows at once.
-- The duplicate releases it before the survivor takes it.
do $$
declare ref text;
begin
  select external_ref into ref from public.expenses
   where id = '92e152d1-ba87-424e-93b3-f93044cbe6aa';

  update public.expenses
     set external_ref = null,
         deleted_at   = now(),
         description  = coalesce(description, '')
                     || ' [duplicate of the 2026-01-17 videographer fee; card settlement '
                     || 'of the same engagement. external_ref moved to the surviving row by 169.]'
   where id = '92e152d1-ba87-424e-93b3-f93044cbe6aa'
     and deleted_at is null;

  update public.expenses
     set external_ref = ref
   where id = '7bbfc00a-e79c-5c94-81a6-d1c963e17a31'
     and external_ref is null
     and ref is not null;
end $$;

update public.actors
   set tax_1099_status = 'exempt',
       tax_1099_note   = 'Card descriptor reads "In *19th & 7th, Inc." — incorporated, '
                      || 'so not 1099-NEC reportable. Confirm entity type with the payee.'
 where deleted_at is null
   and display_name ilike '%19th%'
   and tax_1099_status is null;

-- ---------------------------------------------------------------
-- 2. Gear into Equipment
-- ---------------------------------------------------------------
update public.expenses
   set category = 'Equipment'
 where deleted_at is null
   and ledger = 'come_with'
   and category <> 'Equipment'
   and (   (date = '2026-06-29' and amount = 1513.35)   -- work laptop
        or (date = '2026-06-22' and amount =  400.00)   -- subwoofer
        or (date = '2026-04-26' and amount =   88.74)   -- external hard drive
        or (date = '2025-06-13' and amount =   28.99)   -- cables / accessories
       );

commit;

-- DOWN:
--   clear deleted_at on 92e152d1…, null the external_ref on 7bbfc00a…,
--   restore categories: Production / Production / Operations / Supplies.
