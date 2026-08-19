-- ============================================================
-- COME WITH — 173 the last three gear purchases, and the photographer's name
--
-- 1. 2024 GEAR, CONFIRMED BY KEITH. 170 flagged three charges it would not add
--    on its own because they predate every other Come With expense and could
--    equally have been personal. All three are business:
--
--      2024-08-10  $318.81  Sweetwater  — for the KRK monitors
--      2024-08-12  $318.81  B&H Photo   — for the KRK monitors
--      2024-05-18  $106.60  Best Buy    — equipment accessories, exact item
--                                         not recalled
--
--    All on personal cards, so they are owner capital like the rest of the rig.
--    Equipment 16,093.78 -> 16,838.00; net loss 30,589.98 -> 31,334.20.
--
--    The Best Buy note deliberately records that the item is not remembered
--    rather than inventing a description. An accountant reading "equipment
--    accessories, item not recalled" knows what they are looking at; one reading
--    a confident guess does not.
--
-- 2. THE PHOTOGRAPHER HAS A NAME. The $250 photo shoot on 2026-07-07 was left as
--    'Photographer — name needed (Venmo)' by 167 because the recipient was
--    unknown. It was Tori Mumtaz, who already exists as an actor. Linked, so the
--    payment counts toward her yearly total for 1099 purposes instead of sitting
--    in an unlinked bucket.
-- ============================================================
begin;

insert into public.expenses
  (date, amount, category, vendor, description, ledger,
   cash_source, funded_by, deductible, event_na, external_ref, verified_at)
select v.d::date, v.amt, 'Equipment', v.vendor, v.note, 'come_with',
       'personal', 'owner', true, true, v.ref, now()
  from (values
    ('2024-08-10', 318.81, 'Sweetwater (Benjamin Denen)',
     'Gear for the KRK monitor setup. Confirmed as a business purchase by Keith 2026-08-19.',
     'simplifi:2024-08-10:sweetwater:318.81'),
    ('2024-08-12', 318.81, 'B&H Photo',
     'Gear for the KRK monitor setup. Confirmed as a business purchase by Keith 2026-08-19.',
     'simplifi:2024-08-12:bh-photo:318.81'),
    ('2024-05-18', 106.60, 'Best Buy',
     'Equipment accessories. Exact item not recalled; confirmed as equipment-related by Keith 2026-08-19.',
     'simplifi:2024-05-18:best-buy:106.60')
  ) as v(d, amt, vendor, note, ref)
 where not exists (
   select 1 from public.expenses e
    where e.deleted_at is null and e.date = v.d::date and e.amount = v.amt);

-- Link them to the vendor records that already exist.
update public.expenses x
   set vendor_actor_id = a.id
  from public.actors a
 where a.deleted_at is null
   and a.display_name = x.vendor
   and x.deleted_at is null
   and x.vendor_actor_id is null
   and x.external_ref like 'simplifi:2024-%';

-- ---------------------------------------------------------------
-- 2. Tori Mumtaz
-- ---------------------------------------------------------------
update public.expenses x
   set vendor  = 'Tori Mumtaz',
       vendor_actor_id = a.id,
       description = 'Photo shoot. Paid by Venmo.'
  from public.actors a
 where a.deleted_at is null
   and a.display_name = 'Tori Mumtaz'
   and x.deleted_at is null
   and x.date = date '2026-07-07'
   and x.amount = 250.00
   and x.vendor like 'Photographer%';

insert into public.vendor_aliases (pattern, actor_id, note)
select 'tori mumtaz', a.id, 'photographer — seeded by 173'
  from public.actors a
 where a.deleted_at is null and a.display_name = 'Tori Mumtaz'
on conflict (pattern) do nothing;

commit;

-- DOWN: delete the three expenses by external_ref; restore the photo shoot row's
--   vendor to 'Photographer — name needed (Venmo)' and null its actor.
