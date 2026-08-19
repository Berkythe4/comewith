-- ============================================================
-- COME WITH — 172 We Belong Here festival, June 2026
--
-- Ten charges on 19–21 June were sitting as unattached overhead under
-- Marketing / Networking. They are one trip: We Belong Here, which Keith won
-- tickets to and attended with three people. The Square descriptors confirm it —
-- 'Sq Tce Presents' is TCE Presents, the festival's promoter.
--
--     2 × Uber       19 Jun
--     2 × Uber + Presents  20 Jun
--     5 × on-site Square   21 Jun
--     ----------------------------
--     $598.17 total
--
-- Modelled on the Elements Music & Arts Festival row already in the books —
-- type 'growth', series 'Growth & Networking', not public, no venue. A festival
-- attended to build relationships, not one Come With produced.
--
-- NO REVENUE, and that is a fact rather than a gap: the tickets were won and
-- nothing was sold. Recorded as revenue_confirmed so it stops appearing on the
-- "events with costs but no fee recorded" chase list, which currently counts it
-- against five events genuinely owed money.
--
-- The category stays Marketing / Networking. Attaching an event does not make
-- this a direct event cost in the Come With sense — there was no production to
-- attribute it to — but it does make the trip legible as one line instead of
-- ten unexplained Ubers.
--
-- Guests: Angela Tabone (Dance Infusion) and Zachary Storey already exist as
-- actors. Amanda does not; created with a note asking for her surname rather
-- than inventing one.
-- ============================================================
begin;

insert into public.events
  (slug, name, series, type, status, stage, event_date, end_date,
   is_public, is_content_event, description, notes,
   revenue_confirmed_at)
select
  'we-belong-here-2026', 'We Belong Here Festival', 'Growth & Networking',
  'growth', 'completed', 'wrapped', date '2026-06-19', date '2026-06-21',
  false, false,
  'Festival attended for networking and artist relationships. Tickets won, not bought.',
  'Keith won tickets. Brought Angela Tabone (Dance Infusion), Zachary Storey and '
  || 'Amanda — long-standing supporters. No revenue: nothing was sold and the '
  || 'tickets cost nothing. Costs are travel and on-site spend only.',
  now()
where not exists (select 1 from public.events where slug = 'we-belong-here-2026');

-- Attach the trip's charges. Bounded by date, category and "not already on an
-- event" so it cannot reach past the weekend it belongs to.
update public.expenses x
   set event_id = e.id, event_na = false
  from public.events e
 where e.slug = 'we-belong-here-2026'
   and x.deleted_at is null
   and x.ledger = 'come_with'
   and x.event_id is null
   and x.date between date '2026-06-19' and date '2026-06-21'
   and x.category = 'Marketing / Networking';

-- Amanda has no actor record and no surname on file.
insert into public.actors (display_name, kind, status, notes)
select 'Amanda', 'person', 'active',
       'Attended We Belong Here June 2026 as Keith''s guest. Surname needed.'
where not exists (
  select 1 from public.actors where deleted_at is null and display_name = 'Amanda');

insert into public.event_participants (event_id, actor_id, role, notes)
select e.id, a.id, 'guest', 'Keith''s guest at We Belong Here'
  from public.events e
  join public.actors a on a.deleted_at is null
 where e.slug = 'we-belong-here-2026'
   and a.display_name in ('Angela Tabone', 'Zachary Storey', 'Amanda')
   and not exists (
     select 1 from public.event_participants p
      where p.event_id = e.id and p.actor_id = a.id and p.role = 'guest');

commit;

-- DOWN: null event_id on the 10 charges, delete the participants, delete the
--   event, delete the Amanda actor if it has no other links.
