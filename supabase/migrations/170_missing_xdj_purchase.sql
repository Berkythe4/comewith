-- ============================================================
-- COME WITH — 170 the biggest stolen item was never in the books
--
-- The Pioneer XDJ-AZ heads the stolen-equipment schedule and appears in the
-- accountant memo, but it had NO row in expenses. It was bought on the Discover
-- card on 15 Dec 2024 and Simplifi filed it under 'Entertainment' rather than
-- 'Work Expenses', so the Come With importer never saw it:
--
--     "Dec 15, 2024", Discover, "Sp Pioneer Dj Ca", Entertainment, -3410.93
--
-- Its absence is why equipment totalled $10,651.77 against a theft of
-- $11,837.51 — the books claimed less gear had been bought than was stolen.
--
-- THIS CHANGES HEADLINE FIGURES, deliberately:
--     net loss          28,079.05 -> 30,589.98
--     invested capital  31,630.11 -> 34,141.04
-- (net of the $900 duplicate removed by 169 in the same session)
-- Both were understated by exactly this purchase. It is owner-funded personal
-- card spend like the rest of the rig, so it lands in capital, not the float.
--
-- external_ref is set from the statement identity so a future personal-card
-- import adopts this row instead of inserting a second copy.
--
-- STILL OUTSTANDING, deliberately not added here: three 2024 charges that may or
-- may not be business — B&H $318.81 and Sweetwater $318.81 (both 10-12 Aug 2024)
-- and Best Buy $106.60 (18 May 2024). They predate every other Come With
-- expense, and guessing them in would be inventing a business history that may
-- not exist. Keith's call.
-- ============================================================
begin;

insert into public.expenses
  (date, amount, category, vendor, description, ledger,
   cash_source, funded_by, deductible, event_na, external_ref, verified_at)
select
  date '2024-12-15', 3410.93, 'Equipment', 'Pioneer DJ',
  'Pioneer XDJ-AZ — all-in-one DJ system. Discover card, statement descriptor '
  || '"Sp Pioneer Dj Ca". Miscoded Entertainment in Simplifi, which is why it '
  || 'was missing from the ledger until 2026-08-19. Stolen 16 Aug 2026.',
  'come_with', 'personal', 'owner', true, true,
  'simplifi:2024-12-15:sp-pioneer-dj-ca:3410.93', now()
where not exists (
  select 1 from public.expenses
   where deleted_at is null
     and date = date '2024-12-15'
     and amount = 3410.93
);

commit;

-- DOWN: delete from public.expenses
--        where external_ref = 'simplifi:2024-12-15:sp-pioneer-dj-ca:3410.93';
