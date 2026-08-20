-- ============================================================
-- COME WITH — 180 data health: one place that knows what is not linked yet
--
-- The audit behind 179 found the gaps by hand, with a 200-line query nobody
-- would ever run again. This makes it permanent, self-running and fixable.
--
-- FOUR PIECES:
--   v_data_health          one row per finding, across finance / events / people
--                          / content. Every row names the table, the id, what is
--                          wrong and what fixes it, so a UI can offer the fix.
--   data_health_waivers    the manual override. Anything can be dismissed WITH A
--                          REASON, and it disappears from the list until someone
--                          un-waives it. This is what stops a health check from
--                          becoming noise people learn to scroll past.
--   autolink_data()        does only the links that are provably safe (exact
--                          name, exact email, or a relationship that already
--                          exists), DRY BY DEFAULT, and returns a summary of
--                          what it did or would do.
--   data_health_runs       a row every time either of those runs, with the
--                          summary. Nightly at 07:00 UTC, and on demand from the
--                          dashboard. An automated process that leaves no record
--                          of what it changed is not something you can trust.
--
-- CHECKS ARE WRITTEN TO BE QUIET WHEN THINGS ARE FINE. Three of the hand-audit's
-- checks cried wolf and were rewritten here rather than shipped:
--   * "tickets not linked to a guest" flagged 9 rows that are deliberately LUMP
--     rows (qty 25, 37) holding reconciled Dance Infusion totals. Now only
--     single-seat tickets count.
--   * "guests not linked to an actor" flagged 123 attendees. Attendees are not
--     supposed to be actors; only a guest whose EMAIL already matches an actor
--     is a real missing link. 123 became 1.
--   * "social posts with no event" flagged 23 correct Content Creation slots.
--     179 gave them subject_na, and the check honours it.
-- An alert that is usually wrong trains people to ignore the one time it is
-- right (LEARNINGS §24).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. The manual override
-- ---------------------------------------------------------------
create table if not exists public.data_health_waivers (
  id           uuid primary key default gen_random_uuid(),
  check_key    text not null,
  subject_table text not null,
  subject_id   text not null,
  reason       text not null,
  waived_by    uuid references public.profiles(id),
  waived_at    timestamptz not null default now()
);

create unique index if not exists uq_data_health_waiver
  on public.data_health_waivers(check_key, subject_table, subject_id);

comment on table public.data_health_waivers is
  'A finding somebody has looked at and decided is correct as it stands. Requires '
  'a reason: "dismissed" with no argument is how a check quietly stops meaning '
  'anything. Delete the row to bring the finding back.';

alter table public.data_health_waivers enable row level security;
drop policy if exists "Admins manage data health waivers" on public.data_health_waivers;
create policy "Admins manage data health waivers" on public.data_health_waivers
  for all using (public.is_admin());

-- ---------------------------------------------------------------
-- 2. The run log
-- ---------------------------------------------------------------
create table if not exists public.data_health_runs (
  id          uuid primary key default gen_random_uuid(),
  ran_at      timestamptz not null default now(),
  kind        text not null check (kind in ('audit', 'autolink', 'autolink_dry')),
  source      text not null default 'manual',
  total       integer,
  by_severity jsonb,
  summary     jsonb,
  note        text
);

create index if not exists idx_data_health_runs_at on public.data_health_runs(ran_at desc);

comment on table public.data_health_runs is
  'Every audit sweep and every auto-link, with what it found or changed. This is '
  'the validation record: an automated process nobody can inspect afterwards is '
  'indistinguishable from one that did nothing.';

alter table public.data_health_runs enable row level security;
drop policy if exists "Admins read data health runs" on public.data_health_runs;
create policy "Admins read data health runs" on public.data_health_runs
  for all using (public.is_admin());

-- ---------------------------------------------------------------
-- 3. Every check, as rows
-- ---------------------------------------------------------------
create or replace view public.v_data_health_raw as
-- ============ FINANCE ============
select 'fin.expense_no_payee' as check_key, 'finance' as category, 'high' as severity,
       'Cost with no payee linked' as label,
       'expenses' as subject_table, x.id::text as subject_id,
       coalesce(x.vendor, 'no vendor') || ' · ' || x.amount || ' · ' || x.date as subject_label,
       'Payments cannot be totalled per payee for a 1099 without this' as detail,
       'Link a payee actor, or mark it as having none' as fix_hint
  from public.expenses x
 where x.deleted_at is null and x.vendor_actor_id is null and not x.payee_na
union all
select 'fin.expense_no_event', 'finance', 'high', 'Cost not filed against an event',
       'expenses', x.id::text,
       coalesce(x.vendor, x.category, 'cost') || ' · ' || x.amount || ' · ' || x.date,
       'Sits in overhead until it is filed, so the event looks cheaper than it was',
       'Assign an event, or mark N/A (overhead)'
  from public.expenses x
 where x.deleted_at is null and x.event_id is null and not x.event_na
union all
select 'fin.expense_no_category', 'finance', 'high', 'Cost with no category',
       'expenses', x.id::text,
       coalesce(x.vendor, 'cost') || ' · ' || x.amount || ' · ' || x.date,
       'Lands under Uncategorised on the P&L', 'Pick a category'
  from public.expenses x
 where x.deleted_at is null and (x.category is null or x.category = '')
union all
select 'fin.expense_paid_no_source', 'finance', 'medium', 'Paid cost with no cash source',
       'expenses', x.id::text,
       coalesce(x.vendor, x.category, 'cost') || ' · ' || x.amount,
       'Left out of the cash reserve, so the float reads high',
       'Say which account it left'
  from public.expenses x
 where x.deleted_at is null and x.status = 'paid' and x.cash_source is null
union all
select 'fin.income_no_event', 'finance', 'high', 'Revenue not filed against an event',
       'income', i.id::text,
       coalesce(i.category, 'income') || ' · ' || i.amount || ' · ' || i.date,
       'The event it paid for shows no revenue against its costs',
       'Assign an event, or mark N/A'
  from public.income i
 where i.deleted_at is null and i.event_id is null and not i.event_na
union all
select 'fin.income_no_category', 'finance', 'high', 'Revenue with no stream',
       'income', i.id::text, i.amount || ' · ' || i.date,
       'Cannot be attributed to a revenue stream', 'Pick a revenue stream'
  from public.income i
 where i.deleted_at is null and (i.category is null or i.category = '')
union all
select 'fin.income_recv_no_source', 'finance', 'medium', 'Received revenue with no cash source',
       'income', i.id::text, coalesce(i.category, 'income') || ' · ' || i.amount,
       'Not counted into the cash reserve', 'Say which account it arrived in'
  from public.income i
 where i.deleted_at is null and i.status = 'received' and i.cash_source is null
union all
select 'fin.donation_no_actor', 'finance', 'low', 'Donation with no donor actor',
       'third_party_donations', d.id::text,
       coalesce(d.donor_name, 'anonymous') || ' · ' || d.amount,
       'A repeat donor cannot be seen as one person', 'Link or create the donor actor'
  from public.third_party_donations d
 where d.actor_id is null and coalesce(d.donor_name, '') <> ''
union all
select 'fin.sponsorship_no_actor', 'finance', 'high', 'Sponsorship with no sponsor',
       'sponsorships', s.id::text, coalesce(s.tier, 'sponsorship') || ' · ' || coalesce(s.cash_amount, 0),
       'Sponsor history and renewals cannot be tracked', 'Link the sponsor actor'
  from public.sponsorships s where s.actor_id is null
union all
-- Actionable, not just tidy: money that is late.
select 'fin.payable_overdue', 'finance', 'high', 'Bill past its due date',
       'expenses', p.id::text,
       coalesce(p.payee, 'payee') || ' · ' || p.amount || ' · due ' || p.due_date,
       p.days_overdue || ' days overdue', 'Pay it, or move the due date'
  from public.v_payables p where p.overdue
union all
select 'fin.receivable_stale', 'finance', 'high', 'Money owed to us for over 45 days',
       'income', r.id::text,
       coalesce(r.event_name, r.category, 'invoice') || ' · ' || r.amount,
       r.days_outstanding || ' days since it was recorded', 'Chase it, or settle it'
  from public.v_receivables r where r.days_outstanding > 45
union all
select 'fin.forecast_orphan', 'finance', 'medium', 'Event forecast line with no event',
       'budget_lines', b.id::text, coalesce(b.label, b.category) || ' · ' || b.planned_amount,
       'Counted nowhere, because a forecast is scoped to its event',
       'Attach it to an event or delete it'
  from public.budget_lines b
 where b.scope = 'event' and b.event_id is null and b.deleted_at is null and b.realized_at is null
union all
select 'fin.participant_fee_no_cost', 'finance', 'high', 'Lineup fee never recorded as a cost',
       'event_participants', p.id::text,
       coalesce(a.display_name, 'performer') || ' · ' || p.fee || ' · ' ||
       coalesce((select e.name from public.events e where e.id = p.event_id), 'event'),
       'The fee is promised on the lineup but the event carries no cost for it',
       'Create the payable from the lineup'
  from public.event_participants p
  left join public.actors a on a.id = p.actor_id
 where coalesce(p.fee, 0) > 0
   and not exists (select 1 from public.expenses x
                    where x.deleted_at is null and x.event_id = p.event_id
                      and (x.vendor_actor_id = p.actor_id
                        or lower(coalesce(x.vendor, '')) = lower(coalesce(a.display_name, '~none~'))))
union all
select 'fin.1099_needs_review', 'finance', 'high', 'Payee over $600 with no 1099 decision',
       'actors', coalesce(c.actor_id::text, c.payee), c.payee || ' · ' || c.total_paid || ' in ' || c.tax_year,
       'Over the $600 threshold and nobody has ruled on reportability',
       'Set the payee 1099 status, or link them to an actor first'
  from public.v_contractor_1099 c where c.needs_review
-- ============ EVENTS ============
union all
select 'ev.no_series', 'events', 'high', 'Event with no series',
       'events', e.id::text, e.name || ' · ' || e.event_date,
       'Every KPI matches on series exactly, so this event is invisible to all of them',
       'Set the series'
  from public.events e where e.deleted_at is null and coalesce(e.series, '') = ''
union all
select 'ev.unknown_series', 'events', 'medium', 'Series matches no KPI contract value',
       'events', e.id::text, e.name || ' · ' || e.series,
       'Not counted by any KPI view. Fine for networking events, wrong for anything else',
       'Correct the series, or waive it as intentional'
  from public.events e
 where e.deleted_at is null and coalesce(e.series, '') <> ''
   and e.series not in ('Come With Parties', 'Dance Infusion', 'Come With Production',
                        'Content Creation', 'Bookings', 'Growth & Networking')
union all
select 'ev.no_venue', 'events', 'medium', 'Event with no venue',
       'events', e.id::text, e.name || ' · ' || e.event_date,
       'Venue performance and capacity cannot be compared across events', 'Link a venue'
  from public.events e where e.deleted_at is null and e.venue_id is null
union all
select 'ev.gig_no_owner', 'events', 'medium', 'Booked gig with no host',
       'events', e.id::text, e.name || ' · ' || e.event_date,
       'Bookings roll up by who booked us, and this one rolls up to nobody',
       'Set Host / booked by'
  from public.events e where e.deleted_at is null and e.type = 'gig' and e.owner_actor_id is null
union all
select 'ev.past_no_revenue', 'events', 'high', 'Past event with costs and no revenue',
       'events', m.event_id::text, m.name || ' · ' || m.event_date,
       'Spent ' || m.expenses || ' and recorded nothing coming in',
       'Record the revenue, or confirm it earned nothing'
  from public.v_event_money m where m.missing_revenue
union all
select 'ev.upcoming_no_money', 'events', 'medium', 'Upcoming event with no money against it',
       'events', m.event_id::text, m.name || ' · ' || m.event_date,
       'No revenue, no estimate and no forecast, so it is missing from every projection',
       'Add a forecast line or an expected-revenue estimate'
  from public.v_event_money m
  join public.events e on e.id = m.event_id
 where m.upcoming and m.revenue = 0 and m.forecast_revenue = 0
   and coalesce(e.expected_revenue, 0) = 0 and coalesce(e.status, '') <> 'cancelled'
union all
-- The funnel beacon CANNOT backfill: clicks before a ticket_url exists are lost.
select 'ev.no_ticket_url', 'events', 'high', 'Upcoming public event with no ticket link',
       'events', e.id::text, e.name || ' · ' || e.event_date,
       'Ticket clicks are matched by URL and cannot be backfilled - every click before this is set is lost forever',
       'Set the ticket URL before promotion starts'
  from public.events e
 where e.deleted_at is null and e.is_public and e.event_date > current_date
   and coalesce(e.ticket_url, '') = '' and coalesce(e.status, '') <> 'cancelled'
-- ============ PEOPLE ============
union all
select 'ppl.actor_no_role', 'people', 'medium', 'Actor with no role',
       'actors', a.id::text, a.display_name,
       'Roles drive who appears in every picker, so this actor is hard to find anywhere',
       'Assign a role - it can usually be inferred from what they are linked to'
  from public.actors a
 where a.deleted_at is null
   and not exists (select 1 from public.actor_roles r where r.actor_id = a.id and r.active)
union all
select 'ppl.actor_dupe_email', 'people', 'high', 'Two active actors share an email',
       'actors', a.id::text, a.display_name || ' · ' || a.email,
       'The same person is in the graph twice, so their history is split in half',
       'Merge them from the Actors tab'
  from public.actors a
 where a.deleted_at is null and coalesce(a.email, '') <> ''
   and exists (select 1 from public.actors b
                where b.id <> a.id and b.deleted_at is null
                  and lower(b.email) = lower(a.email))
union all
select 'ppl.guest_is_an_actor', 'people', 'medium', 'Guest whose email already matches an actor',
       'guests', g.id::text, coalesce(g.full_name, g.email),
       'The same person is a guest here and an actor there, with no link between them',
       'Link the guest to the actor'
  from public.guests g
 where g.deleted_at is null and g.actor_id is null and coalesce(g.email, '') <> ''
   and exists (select 1 from public.actors a
                where a.deleted_at is null and lower(a.email) = lower(g.email))
union all
select 'ppl.subscriber_no_brand', 'people', 'high', 'Subscriber in no brand segment',
       'subscribers', s.id::text, s.email,
       'Campaigns target brand segments, so this person receives nothing',
       'Add come_with or dance_infusion'
  from public.subscribers s
 where s.unsubscribed_at is null
   and not exists (select 1 from public.subscriber_segments g
                    where g.subscriber_id = s.id and g.segment in ('come_with', 'dance_infusion'))
union all
select 'ppl.ticket_no_guest', 'people', 'low', 'Single ticket not linked to a guest',
       'ticketing', t.id::text,
       coalesce(t.ticket_type, 'ticket') || ' · ' ||
       coalesce((select e.name from public.events e where e.id = t.event_id), 'event'),
       'Attendance and repeat-customer history cannot follow this ticket',
       'Link the buyer'
  from public.ticketing t where t.guest_id is null and coalesce(t.quantity, 1) = 1
-- ============ CONTENT / OPS ============
union all
select 'ops.post_no_subject', 'content', 'low', 'Social post about no event and no episode',
       'social_posts', p.id::text, coalesce(p.title, 'untitled') || ' · ' || coalesce(p.stage, ''),
       'Content cannot be credited to what it was promoting',
       'Link an event or episode, or mark it evergreen'
  from public.social_posts p
 where p.deleted_at is null and p.event_id is null and p.station_id is null and not p.subject_na
union all
select 'ops.photo_no_credit', 'content', 'low', 'Photo with no photographer credited',
       'event_photos', f.id::text,
       coalesce(f.shoot_label, (select e.name from public.events e where e.id = f.event_id), 'photo'),
       'Cannot produce a credit line, and the photographer cannot see their own portfolio',
       'Credit the photographer, in bulk from the Photos tab'
  from public.event_photos f where f.photographer_actor_id is null
union all
select 'ops.task_orphan', 'content', 'low', 'Open task attached to nothing',
       'tasks', t.id::text, t.title,
       'Does not appear on any event, episode or meeting', 'Attach it, or leave it standalone'
  from public.tasks t
 where t.deleted_at is null and coalesce(t.status, '') <> 'done'
   and t.event_id is null and t.station_id is null and t.post_id is null and t.meeting_id is null
union all
select 'ops.equipment_no_event', 'content', 'low', 'Equipment use with no event',
       'equipment_usage', u.id::text, coalesce(u.purpose, 'usage'),
       'Gear utilisation cannot be attributed', 'Attach the event'
  from public.equipment_usage u where u.event_id is null;

revoke select on public.v_data_health_raw from anon;

-- The list people actually read: waived findings removed.
create or replace view public.v_data_health as
select r.*
  from public.v_data_health_raw r
 where not exists (select 1 from public.data_health_waivers w
                    where w.check_key = r.check_key
                      and w.subject_table = r.subject_table
                      and w.subject_id = r.subject_id);

revoke select on public.v_data_health from anon;

create or replace view public.v_data_health_summary as
select category, check_key, severity, label,
       min(detail)   as detail,
       min(fix_hint) as fix_hint,
       count(*)      as open_count,
       (select count(*) from public.data_health_waivers w where w.check_key = h.check_key) as waived_count
  from public.v_data_health h
 group by category, check_key, severity, label;

revoke select on public.v_data_health_summary from anon;

commit;

-- DOWN: drop view public.v_data_health_summary, public.v_data_health,
--   public.v_data_health_raw; drop table public.data_health_runs,
--   public.data_health_waivers;
