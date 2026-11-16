-- =============================================================================
-- 248_hulaween_merge_consulting_cash.sql - Keith, 2026-11-16
-- =============================================================================
--   1. ONE Hulaween. 245 created `hulaween-2026` without noticing the original
--      `hulaween` event (July; 10/22-25, Growth & Networking) that already held
--      20 card-imported charges, 4 lineup rows and a task. Keith: the dates are
--      10/12-10/21. The 245 event already carries those dates, Consulting,
--      JSD Drip and both CWC invoices, so it SURVIVES; the original's rows move
--      onto it (lineup roles merged) and the original is soft-deleted.
--   2. E-ZPass rebill ($25) and Morphis Art ($21.20) are "left off" Hulaween
--      (Keith): detached from the event, left in the Come With books as overhead.
--   3. Consulting expenses paid on Keith's personal card are CONSULTING's money
--      (Consulting is Keith, a sole proprietorship), not owner capital put into
--      Come With: funded_by owner -> business. v_capital has no ledger filter,
--      so leaving them 'owner' would show them as Come With owing Keith.
--   4. McDonald's $31.58 + Subway $40.35 (Hulaween meals, now Consulting) were
--      paid from Come With's bank: Consulting owed Come With $71.93. Keith:
--      settle it against Come With expenses he paid on his personal card. Two
--      sum to exactly $71.93 (Beatport 2025-07-26 $5.96, Sq Tce Presents
--      2026-06-21 $65.97); they are marked reimbursed with the reason recorded.
--   5. v_cash_position: a CONSULTING expense paid from 'bank' left Come With's
--      bank (Consulting has no business account), so it counts in Come With's
--      outflow. Without this the $71.93 vanished from the cash reserve the
--      moment the meals moved books. Same columns; only the outflow predicate.
-- =============================================================================

begin;

do $$
declare
  orig uuid := 'ffc2391c-c102-4150-a964-205bbe1c1e2c';
  keep uuid := 'b12ea8b3-5d60-4e57-ad0c-01cbd53aaf2c';
  n int;
begin
  if not exists (select 1 from public.events where id = orig and deleted_at is null) then raise exception 'original Hulaween already gone'; end if;

  -- 2. off Hulaween (before the move, so they are not carried over)
  update public.expenses set event_id = null, event_na = true,
         description = coalesce(description || ' ', '') || '[Taken off Hulaween 2026-11-16 per Keith]'
   where event_id = orig and deleted_at is null and vendor in ('E-z*passny Rebill', 'Sp Morphis Art Llc');
  get diagnostics n = row_count; if n <> 2 then raise exception 'expected 2 rows off Hulaween, got %', n; end if;

  -- 1. move everything else
  update public.expenses set event_id = keep where event_id = orig;
  get diagnostics n = row_count; if n <> 18 then raise exception 'expected 18 expenses to move (16 live + 2 deleted? got %)', n; end if;
  update public.tasks set event_id = keep where event_id = orig;
  -- lineup: union the roles where the actor is on both, move the rest
  update public.event_participants k set roles = (select array(select distinct unnest(k.roles || o.roles))), notes = coalesce(k.notes, o.notes)
    from public.event_participants o where k.event_id = keep and o.event_id = orig and o.actor_id = k.actor_id;
  delete from public.event_participants o where o.event_id = orig
     and exists (select 1 from public.event_participants k where k.event_id = keep and k.actor_id = o.actor_id);
  update public.event_participants set event_id = keep where event_id = orig;
  update public.events set deleted_at = now(),
         notes = coalesce(notes || ' ', '') || '[Merged into hulaween-2026 (' || keep || ') on 2026-11-16: duplicate created by 245.]'
   where id = orig;
  update public.events set notes = 'Come With Consulting work for JSD Drip, Oct 12-21. $1,350 paid in kind (joggers rev-share buy-in, CWC-2026-0001). Travel billed back on CWC-2026-0002. Merged with the July "Hulaween" event on 2026-11-16.'
   where id = keep;

  -- 3. Consulting's own money
  update public.expenses set funded_by = 'business'
   where ledger = 'consulting' and cash_source = 'personal' and funded_by = 'owner' and deleted_at is null;

  -- 4. the $71.93 debt, settled by offset
  if (select sum(amount) from public.expenses where id in ('969fbd66-1905-5d90-97ba-24dbe1b52e46', '0781ca90-6c1e-4748-ace2-a3506bcdde5f')
        and ledger = 'come_with' and funded_by = 'owner' and reimbursed_at is null and deleted_at is null) <> 71.93
    then raise exception 'offset rows are not the expected unreimbursed $71.93'; end if;
  update public.expenses set reimbursed_at = now(),
         description = coalesce(description || ' ', '') || '[Reimbursed 2026-11-16 by offset: Come With''s bank paid $71.93 of Come With Consulting''s Hulaween meals (McDonald''s 10/21 $31.58 + Subway 10/22 $40.35), settled against these personal-card Come With costs.]'
   where id in ('969fbd66-1905-5d90-97ba-24dbe1b52e46', '0781ca90-6c1e-4748-ace2-a3506bcdde5f');
  update public.expenses set description = coalesce(description || ' ', '') || '[Paid from Come With''s bank for Consulting; repaid by offset 2026-11-16 (Beatport 2025-07-26 + Sq Tce Presents 2026-06-21).]'
   where event_id = keep and ledger = 'consulting' and cash_source = 'bank' and deleted_at is null;
  get diagnostics n = row_count; if n <> 2 then raise exception 'expected the 2 bank-paid meals, got %', n; end if;
end $$;

-- 5. Come With's cash includes Consulting costs it paid from its own bank
create or replace view public.v_cash_position as
 WITH cap AS (
         SELECT COALESCE(sum(capital_contributions.amount), (0)::numeric) AS amt
           FROM capital_contributions
          WHERE ((capital_contributions.deleted_at IS NULL) AND (capital_contributions.kind = 'cash'::text))
        ), inflow AS (
         SELECT COALESCE(sum(income.amount), (0)::numeric) AS amt
           FROM income
          WHERE ((income.deleted_at IS NULL) AND (income.ledger = 'come_with'::text) AND (income.status = 'received'::text) AND (income.cash_source = ANY (ARRAY['paypal'::text, 'bank'::text])))
        ), outflow AS (
         SELECT COALESCE(sum(expenses.amount), (0)::numeric) AS amt
           FROM expenses
          WHERE ((expenses.deleted_at IS NULL) AND (expenses.status = 'paid'::text)
                 AND ((expenses.ledger = 'come_with'::text AND expenses.cash_source = ANY (ARRAY['paypal'::text, 'bank'::text]))
                      -- 248: Consulting has no business account; 'bank' on a Consulting cost is Come With's bank
                      OR (expenses.ledger = 'consulting'::text AND expenses.cash_source = 'bank'::text)))
        ), unknown_src AS (
         SELECT count(*) AS n,
            COALESCE(sum(expenses.amount), (0)::numeric) AS amt
           FROM expenses
          WHERE ((expenses.deleted_at IS NULL) AND (expenses.ledger = 'come_with'::text) AND (expenses.status = 'paid'::text) AND (expenses.cash_source IS NULL))
        ), owed AS (
         SELECT COALESCE(sum(v_receivables.amount), (0)::numeric) AS amt,
            count(*) AS n
           FROM v_receivables
          WHERE (v_receivables.ledger = 'come_with'::text)
        ), owing AS (
         SELECT COALESCE(sum(v_payables.amount), (0)::numeric) AS amt,
            count(*) AS n
           FROM v_payables
          WHERE (v_payables.ledger = 'come_with'::text)
        )
 SELECT cap.amt AS capital_in,
    inflow.amt AS revenue_in,
    outflow.amt AS spent_from_float,
    ((cap.amt + inflow.amt) - outflow.amt) AS cash_reserve,
    unknown_src.n AS unknown_source_rows,
    unknown_src.amt AS unknown_source_amount,
    owed.amt AS owed_to_us,
    owed.n AS owed_count,
    owing.amt AS owed_by_us,
    owing.n AS owing_count,
    ((((cap.amt + inflow.amt) - outflow.amt) - owing.amt) + owed.amt) AS cash_after_commitments
   FROM cap, inflow, outflow, unknown_src, owed, owing;

commit;
