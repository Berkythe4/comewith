-- ============================================================
-- COME WITH — 166 the 1099 list is per payee, not per category
--
-- 165 built v_contractor_1099 off category='Contractors'. That is the wrong
-- basis and it under-reported real money:
--
--   Janelle Sochet        $700 by that view, $900 actually paid
--                         ($200 sat in 'Marketing', tied to an event)
--   19th & 7th (McManus)  $900 by that view, $1,800 actually paid
--                         ($900 sat in 'Production')
--
-- The $600 reporting threshold is measured on total payments to a payee in a
-- calendar year for services. It does not care which internal bucket we filed
-- them under, so neither can this view. Both payees crossed the threshold on
-- the true figures and would have been reported short.
--
-- WHAT THIS VIEW CANNOT DECIDE. Whether a 1099-NEC is actually owed turns on
-- facts the ledger does not hold: is the payee a corporation, was this services
-- or goods, was it a reimbursement. Beatport at $939 is software and obviously
-- not reportable; 19th & 7th is a production company and might be incorporated.
-- Guessing from the category is what produced the wrong numbers above.
--
-- So the decision is stored, not inferred: actors.tax_1099_status, set by a
-- person. Anything unset shows as 'undecided' and stays on the list. An
-- incomplete list that says so is worth more than a confident wrong one.
--
-- The obvious goods and software vendors are pre-marked exempt below, because
-- making someone hand-classify Amazon is how a review gets abandoned halfway.
-- ============================================================
begin;

alter table public.actors
  add column if not exists tax_1099_status text
    check (tax_1099_status in ('due', 'exempt')),
  add column if not exists tax_1099_note text;

comment on column public.actors.tax_1099_status is
  'Human decision on 1099-NEC reportability. null = not yet reviewed. '
  'exempt covers corporations, goods vendors, and reimbursements.';

-- Pre-classify the payees where there is no judgement to make: retailers,
-- software subscriptions, and gear. Names matched loosely because these arrive
-- from statement descriptors.
update public.actors a
   set tax_1099_status = 'exempt',
       tax_1099_note   = coalesce(tax_1099_note, 'goods or software vendor, not services')
 where a.deleted_at is null
   and a.tax_1099_status is null
   and (a.display_name ilike any (array[
         '%sweetwater%', '%best buy%', '%amazon%', '%beatport%', '%anthropic%',
         '%b&h%', '%krk%', '%meta%', '%google%', '%apple%', '%adobe%',
         '%squarespace%', '%netlify%', '%supabase%', '%openai%', '%canva%'
       ]));

-- ---------------------------------------------------------------
-- Column set changes meaning, so replace rather than append.
-- ---------------------------------------------------------------
drop view if exists public.v_contractor_1099;

create view public.v_contractor_1099 as
select
  coalesce(a.display_name, x.vendor)            as payee,
  x.vendor_actor_id                             as actor_id,
  extract(year from x.date)::int                as tax_year,
  count(*)                                      as payments,
  round(sum(x.amount), 2)                       as total_paid,
  min(x.date)                                   as first_payment,
  max(x.date)                                   as last_payment,
  string_agg(distinct x.category, ', ' order by x.category) as categories,
  (sum(x.amount) >= 600)                        as over_threshold,
  round(greatest(600 - sum(x.amount), 0), 2)    as headroom,
  coalesce(a.tax_1099_status, 'undecided')      as status,
  a.tax_1099_note                               as note,
  -- The working list: crossed the threshold and nobody has ruled on it yet.
  (sum(x.amount) >= 600 and a.tax_1099_status is null) as needs_review
from public.expenses x
left join public.actors a on a.id = x.vendor_actor_id
where x.deleted_at is null
  and x.ledger = 'come_with'
group by 1, 2, 3, a.tax_1099_status, a.tax_1099_note;

revoke select on public.v_contractor_1099 from anon;

commit;

-- DOWN:
--   drop view public.v_contractor_1099;
--   alter table public.actors drop column tax_1099_status, drop column tax_1099_note;
