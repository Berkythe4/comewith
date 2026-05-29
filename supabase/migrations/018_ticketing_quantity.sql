-- ============================================================
-- COME WITH — 018 TICKETING QUANTITY  (additive)
-- Per-line quantity so Log Event writes a few rows (one per price
-- tier) instead of one row per attendee. Existing/importer rows
-- (one row per ticket) default to 1, preserving their semantics.
--
-- Convention: amount_paid = LINE TOTAL (unit price x quantity).
-- So sum(amount_paid) = revenue and sum(quantity) = ticket count
-- stay coherent for both tier rows and per-ticket rows.
-- ============================================================
begin;

alter table public.ticketing
  add column if not exists quantity integer not null default 1 check (quantity > 0);

-- tickets_sold must count tickets, not rows. Update the single source
-- of truth (v_event_summary); v_kpi_event_financials / v_kpi_parties
-- reuse it. Output columns/types unchanged -> CREATE OR REPLACE is safe.
create or replace view public.v_event_summary as
select
  e.id as event_id, e.slug, e.name, e.event_date, e.series, e.status, e.venue_id,
  coalesce(rev.revenue,0)  as revenue,
  coalesce(exp.expenses,0) as expenses,
  coalesce(rev.revenue,0) - coalesce(exp.expenses,0) as net,
  coalesce(spn.sponsor_count,0) as sponsor_count,
  coalesce(spn.sponsor_cash,0)  as sponsor_cash,
  coalesce(tkt.tickets_sold,0)  as tickets_sold,
  coalesce(tkt.ticket_revenue,0) as ticket_revenue,
  e.total_attendance,
  coalesce(dn.third_party_total,0) as third_party_donations
from public.events e
left join lateral (select sum(amount) as revenue  from public.income   where event_id=e.id and deleted_at is null) rev on true
left join lateral (select sum(amount) as expenses from public.expenses where event_id=e.id and deleted_at is null) exp on true
left join lateral (select count(*) as sponsor_count, sum(cash_amount+in_kind_value) as sponsor_cash
                   from public.sponsorships where event_id=e.id and status != 'cancelled') spn on true
left join lateral (select sum(coalesce(quantity,1)) as tickets_sold, sum(amount_paid) as ticket_revenue
                   from public.ticketing where event_id=e.id) tkt on true
left join lateral (select sum(amount) as third_party_total from public.third_party_donations where event_id=e.id) dn on true
where e.deleted_at is null;

-- v_event_summary had anon SELECT revoked in 015; CREATE OR REPLACE
-- preserves grants, but re-assert to be certain it stays non-public.
revoke select on public.v_event_summary from anon;

commit;
