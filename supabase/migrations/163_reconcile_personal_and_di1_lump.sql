-- ============================================================
-- COME WITH — 163 everything not PayPal or Bluevine is personal; DI#1 itemised
--
-- 1. CASH SOURCE, FINISHED. The full PayPal and Bluevine histories are now
--    loaded, so by elimination every remaining Come With charge was paid on the
--    personal card. Setting it explicitly makes the $5,000 float exact instead of
--    conservative, and marks the spend as owner-funded capital, which is what it
--    is.
--
--    Scoped to ledger='come_with' ON PURPOSE. Dance Infusion banks separately, so
--    sweeping its rows into "Keith paid this personally" would inflate the capital
--    he has put into Come With with money that went to a charity instead.
--
-- 2. DI#1 WITHOUT DOUBLE COUNTING. The four 2025 Signal charges ($862.77) are bar
--    minimum spend for Dance Infusion #1. They are ALREADY inside DI#1's books —
--    as part of a single $1,635 line reading "DI#1 event costs (founder-paid)",
--    which is also why the $1,800 donation from Keith exists: his personal spend
--    IS that donation.
--
--    So attaching the Signal rows on their own would count the same money twice
--    and push cost-to-raise from $0.61 to $0.90 on a figure published to the
--    public. Instead the lump is REDUCED by exactly what is being itemised:
--
--        1,635.00 - 862.77 = 772.23   (remaining, still unitemised)
--        60 + 105 + 772.23 + 862.77   = 1,800.00, unchanged
--
--    Every published metric holds: raised $2,942.50, net $1,142.50, cost to raise
--    $0.61. A lump becomes real transactions and nothing moves.
--
-- 3. v_capital was summing owner-funded spend across BOTH ledgers, so DI costs
--    would have counted as capital invested in Come With. Now scoped.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Itemise DI#1 before touching cash source, so the Signal rows are already
--    on the DI ledger when the come_with sweep runs and cannot be caught by it.
-- ---------------------------------------------------------------
do $$
declare
  di1  uuid;
  moved numeric;
begin
  select id into di1 from public.events
   where name = 'Dance Infusion #1' and deleted_at is null;
  if di1 is null then
    raise notice 'Dance Infusion #1 not found — skipping itemisation';
    return;
  end if;

  select coalesce(sum(amount), 0) into moved
    from public.expenses
   where deleted_at is null and event_id is null
     and vendor ilike '%signal%' and date < '2026-01-01';

  if moved = 0 then
    raise notice 'no unlinked 2025 Signal charges — already done?';
    return;
  end if;

  -- Attach them to the event, on the Dance Infusion books.
  update public.expenses
     set event_id = di1, event_na = false, ledger = 'dance_infusion',
         category = 'Venue', cash_source = 'personal',
         description = coalesce(description, 'Bar minimum spend — Dance Infusion #1')
   where deleted_at is null and event_id is null
     and vendor ilike '%signal%' and date < '2026-01-01';

  -- And take the same amount out of the lump they were hiding inside.
  update public.expenses
     set amount = amount - moved,
         description = 'DI#1 event costs (founder-paid), remaining unitemised. '
                    || 'Reduced by ' || to_char(moved, 'FM999999.00')
                    || ' on 2026-08-19 when the Signal bar-minimum charges were itemised.'
   where deleted_at is null and event_id = di1
     and description like 'DI#1 event costs (founder-paid)%';

  raise notice 'itemised % from the DI#1 lump', moved;
end $$;

-- ---------------------------------------------------------------
-- 2. Everything else on the Come With books was the personal card
-- ---------------------------------------------------------------
update public.expenses
   set cash_source = 'personal',
       funded_by = 'owner'
 where deleted_at is null
   and ledger = 'come_with'
   and cash_source is null;

-- Dance Infusion's own spend is not Come With capital. Where it has no source,
-- say "other" rather than claiming a pot it never came from.
update public.expenses
   set cash_source = 'other'
 where deleted_at is null
   and ledger = 'dance_infusion'
   and cash_source is null;

-- ---------------------------------------------------------------
-- 3. Capital means capital in COME WITH
-- ---------------------------------------------------------------
create or replace view public.v_capital as
with contrib as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions where deleted_at is null
), personal as (
  select coalesce(sum(amount), 0) as amt,
         coalesce(sum(amount) filter (where reimbursed_at is not null), 0) as repaid
    from public.expenses
   where deleted_at is null and funded_by = 'owner' and ledger = 'come_with'
)
select
  contrib.amt                                  as contributed,
  personal.amt                                 as personally_paid,
  personal.repaid                              as reimbursed,
  contrib.amt + personal.amt                   as invested_gross,
  contrib.amt + personal.amt - personal.repaid as invested_net,
  personal.amt - personal.repaid               as outstanding_reimbursable
from contrib, personal;

revoke select on public.v_capital from anon;

commit;

-- DOWN: restore from backup — the lump reduction is not reconstructible here.
