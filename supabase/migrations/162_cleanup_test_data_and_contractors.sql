-- ============================================================
-- COME WITH — 162 remove the test events, and code vendor retainers properly
--
-- 1. TEST DATA. 'Test' and 'ZZ delete me' were soft-deleted long ago but kept two
--    ticketing rows worth $550, which 161 stopped counting but did not remove.
--    Leaving orphaned rows under deleted events is how the next person loses an
--    afternoon. Tickets go first, then the events themselves.
--
-- 2. VENDOR RETAINERS. Two payments were sitting in 'Operations', which says
--    nothing:
--      Janelle Sochet  $900 over 4 payments — ongoing marketing work
--      Henry Zaradich  $100 — reimbursing his Claude subscription so he can keep
--                      building the site
--    Both are the same thing: paying a person for ongoing work that belongs to no
--    single event. That is 'Contractors' — a category that already exists here.
--
--    Henry's is a tool reimbursement, so 'Software' is tempting. It is the wrong
--    call: the money went to a contractor as part of what it costs to keep him
--    working, and coding it as Software would put a person's retainer in the same
--    bucket as Ableton. What you are buying is his time.
--
--    Both stay INDIRECT (no event) — a retainer is overhead, which is exactly
--    what makes it a commitment rather than a one-off.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Test data out
-- ---------------------------------------------------------------
delete from public.ticketing t
 using public.events e
 where e.id = t.event_id and e.name in ('Test', 'ZZ delete me');

-- Any other financial residue under those events, before the events go.
delete from public.income i
 using public.events e
 where e.id = i.event_id and e.name in ('Test', 'ZZ delete me');

delete from public.expenses x
 using public.events e
 where e.id = x.event_id and e.name in ('Test', 'ZZ delete me');

delete from public.events where name in ('Test', 'ZZ delete me');

-- ---------------------------------------------------------------
-- 2. Retainers are contractor spend, not "Operations"
-- ---------------------------------------------------------------
update public.expenses
   set category = 'Contractors'
 where deleted_at is null
   and event_id is null
   and (vendor ilike '%janelle%' or vendor ilike '%sochet%' or vendor ilike '%henry%')
   and coalesce(category, '') <> 'Contractors';

-- Make it stick for the next import rather than needing this again.
insert into public.vendor_aliases (pattern, actor_id, note)
select v.pattern, a.id, 'retainer — coded to Contractors by 162'
  from (values ('janelle', 'Janelle Sochet'), ('sochetjanel', 'Janelle Sochet'), ('henry', 'Henry'))
       as v(pattern, actor_name)
  join public.actors a on a.deleted_at is null and a.display_name = v.actor_name
on conflict (pattern) do update set note = excluded.note;

commit;

-- DOWN: the deletions are not reversible from this file — restore from backup.
--   Category changes: set category='Operations' for the same rows.
