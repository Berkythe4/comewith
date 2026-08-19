-- ============================================================
-- COME WITH — 171 give Pioneer a vendor record; make "unlinked" say so
--
-- 1. The XDJ purchase added by 170 named its vendor only as free text, so it
--    surfaced on the 1099 review list as an undecided $3,410.93 payee. It is
--    Pioneer DJ, a manufacturer — goods, never reportable. Given a proper actor
--    record so the exemption has somewhere to live, plus an alias so future
--    Pioneer charges link themselves.
--
-- 2. A GAP THE 166 DESIGN LEFT. Reportability is stored on the actor, so a payee
--    that exists only as vendor text has nowhere to record a decision and sits
--    on the review list permanently, with no action available that would clear
--    it. The view now separates the two cases:
--
--        'undecided'  — there is an actor, nobody has ruled on it
--        'no vendor'  — no actor at all; link one first, then decide
--
--    Different problems, different fixes, so they should not share a label.
--
-- 3. 'Best Buy' was recorded as kind='person'. Harmless to the 1099 answer since
--    it is already exempt, but it is a shop, and actor kind is what the vendor
--    pickers group on.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Pioneer DJ as a vendor
-- ---------------------------------------------------------------
insert into public.actors (display_name, kind, status, tax_1099_status, tax_1099_note)
select 'Pioneer DJ', 'org', 'active', 'exempt',
       'Equipment manufacturer — goods, not services. Never 1099 reportable.'
where not exists (
  select 1 from public.actors where deleted_at is null and display_name = 'Pioneer DJ');

update public.expenses x
   set vendor_actor_id = a.id
  from public.actors a
 where a.display_name = 'Pioneer DJ' and a.deleted_at is null
   and x.deleted_at is null
   and x.vendor_actor_id is null
   and x.vendor ilike '%pioneer%';

insert into public.vendor_aliases (pattern, actor_id, note)
select 'pioneer dj', a.id, 'manufacturer — seeded by 171'
  from public.actors a
 where a.display_name = 'Pioneer DJ' and a.deleted_at is null
on conflict (pattern) do nothing;

-- ---------------------------------------------------------------
-- 2. Tell the two "not decided" cases apart
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
  case when x.vendor_actor_id is null then 'no vendor'
       else coalesce(a.tax_1099_status, 'undecided') end as status,
  a.tax_1099_note                               as note,
  (sum(x.amount) >= 600
   and (x.vendor_actor_id is null or a.tax_1099_status is null)) as needs_review
from public.expenses x
left join public.actors a on a.id = x.vendor_actor_id
where x.deleted_at is null
  and x.ledger = 'come_with'
group by 1, 2, 3, a.tax_1099_status, a.tax_1099_note;

revoke select on public.v_contractor_1099 from anon;

-- ---------------------------------------------------------------
-- 3. A shop is not a person
-- ---------------------------------------------------------------
update public.actors set kind = 'org'
 where deleted_at is null and display_name = 'Best Buy' and kind = 'person';

commit;

-- DOWN: restore the 166/167 view definition; delete the Pioneer DJ actor and
--   alias; set Best Buy kind back to 'person'.
