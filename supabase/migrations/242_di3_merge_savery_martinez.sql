-- =============================================================================
-- 242_di3_merge_savery_martinez.sql - Keith, 2026-11-11 (DI #3)
-- =============================================================================
--   Savery/Martinez is ONE Champion sponsorship ($500), not two cards. Arnaldo
--   Martinez's $250 half (pending, 0 comps, 0 drinks) folds into Patrick
--   Savery's sponsorship, which already holds the 4 comps, the 4-ticket drink
--   package and all three named comps (Patrick, Arnaldo, Alice).
--     * Savery card: cash 250 -> 500, status paid -> confirmed (only half is in;
--       'paid' at $500 would overstate it), notes record who paid what.
--     * Arnaldo's sponsorship row is deleted. Nothing else references it
--       (gea.sponsorship_id is the only FK column; none point at it).
--   His actor record stays. Pledged total for the event is unchanged ($500).
-- =============================================================================

begin;

do $$
declare
  sp_sav constant uuid := '5e998046-eb9f-45c0-b6e3-ac4699aa58b6';
  sp_am  constant uuid := '110fc03f-a016-4c4f-88f9-e77518a90a56';
  pledged_before numeric; n int;
begin
  select sum(cash_amount) into pledged_before from public.sponsorships
   where event_id = (select event_id from public.sponsorships where id = sp_sav) and status <> 'cancelled';

  if exists (select 1 from public.guest_event_attendance where sponsorship_id = sp_am) then
    raise exception 'door-list rows still point at Arnaldo''s sponsorship';
  end if;

  delete from public.sponsorships
   where id = sp_am and cash_amount = 250 and entry_tickets = 0 and drink_tickets = 0 and status = 'pending';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Arnaldo sponsorship not as expected (already merged?)'; end if;

  update public.sponsorships
     set cash_amount = 500, status = 'confirmed',
         notes = 'Savery/Martinez Champion, $500 total (merged 2026-11-11). Patrick Savery: $250 paid via Zeffy 10/26/2026 (logged there as a donation; reclassified 2026-11-04). Dr. Arnaldo Martinez: $250 pending - mark paid when it lands. 4 comps (Patrick, Arnaldo, Alice + 1 open) and a 4-ticket drink package, currently all on Patrick.'
   where id = sp_sav and cash_amount = 250 and status = 'paid';
  get diagnostics n = row_count; if n <> 1 then raise exception 'Savery sponsorship not as expected'; end if;

  if (select sum(cash_amount) from public.sponsorships
       where event_id = (select event_id from public.sponsorships where id = sp_sav) and status <> 'cancelled') <> pledged_before
    then raise exception 'pledged total moved'; end if;
  if (select count(*) from public.guest_event_attendance where sponsorship_id = sp_sav) <> 3 then raise exception 'Savery comps <> 3 named'; end if;
end $$;

commit;
