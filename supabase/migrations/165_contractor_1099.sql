-- ============================================================
-- COME WITH — 165 track contractor pay against the 1099 threshold
--
-- Martin, Henry and Janelle can each earn 5% by sweat equity over two years, but
-- TODAY they are contractors. That has a filing consequence nobody is currently
-- watching: pay an unincorporated contractor $600 or more in a calendar year and
-- a 1099-NEC is due by 31 January.
--
-- Two people are already over it for 2026 — Janelle Sochet ($900) and 19th & 7th
-- Productions ($900) — and the deadline is months away, which is exactly how it
-- gets missed. This makes the threshold something the dashboard can see rather
-- than something someone has to remember in January.
--
-- Deliberately NOT automatic: whether a 1099 is actually required depends on the
-- payee's entity type (corporations are generally exempt) and on whether the
-- payment was fee or reimbursement. Henry's $100 is a reimbursement for a Claude
-- subscription, which is a different thing from a fee. The view reports the
-- money; a human decides the form.
-- ============================================================
begin;

create or replace view public.v_contractor_1099 as
select
  coalesce(a.display_name, x.vendor)                as payee,
  x.vendor_actor_id                                 as actor_id,
  extract(year from x.date)::int                    as tax_year,
  count(*)                                          as payments,
  round(sum(x.amount), 2)                           as total_paid,
  min(x.date)                                       as first_payment,
  max(x.date)                                       as last_payment,
  (sum(x.amount) >= 600)                            as over_threshold,
  round(greatest(600 - sum(x.amount), 0), 2)        as headroom
from public.expenses x
left join public.actors a on a.id = x.vendor_actor_id
where x.deleted_at is null
  and x.ledger = 'come_with'
  and x.category = 'Contractors'
group by 1, 2, 3;

revoke select on public.v_contractor_1099 from anon;

comment on view public.v_contractor_1099 is
  'Contractor pay per person per calendar year against the $600 1099-NEC '
  'threshold. Reports money only — whether a form is required depends on the '
  'payee''s entity type and on fee vs reimbursement.';

commit;

-- DOWN: drop view if exists public.v_contractor_1099;
