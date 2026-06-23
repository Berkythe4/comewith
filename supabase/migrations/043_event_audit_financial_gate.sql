-- =============================================================================
-- 043_event_audit_financial_gate.sql
-- Event-scoped financial publish gate + master/sub financial split.
--
-- ┌──────────────────────────────────────────────────────────────────────────┐
-- │  NOT YET APPLIED TO PROD.  Financial carve-out per Keith: "no changes to   │
-- │  financials should happen" in the applied run. This is the staged design.  │
-- │  Apply only after review + anon-revoke verification (the 016/017->019      │
-- │  regression class lives exactly here).                                     │
-- └──────────────────────────────────────────────────────────────────────────┘
--
-- Model:
--   * events.audited (bool) + audited_at + audited_by. Only master_admin may
--     flip it. While FALSE: that event's financials are master_admin-only.
--     When master flips it TRUE: that event's FINAL financials become visible
--     to all staff (full transparency on Keith's trigger).
--   * Company-level rows (income/expenses with event_id IS NULL = overhead /
--     general income) are master_admin-only PERMANENTLY — no per-event flag can
--     release them. Enforced by the IS NOT NULL + audited join below.
--   * Strategy rollups (v_kpi_dashboard and the aggregate KPI views) stay
--     master_admin-only for now — released separately, not by this flag.
--
-- Why the gate MUST live in the view layer (from the audit): v_event_summary,
-- v_kpi_event_financials, v_kpi_parties, v_kpi_dance_infusion, v_kpi_dashboard
-- were created WITHOUT security_invoker (015/022) so they run as owner and
-- BYPASS base-table RLS. Two-part fix below:
--   (A) Gate the financial COLUMNS inside each view with a CASE on
--       is_master_admin() OR e.audited  -> NULL for staff on unaudited events,
--       so an unaudited event simply shows blank financials, never a leak.
--   (B) Also set security_invoker=true and add audited-aware base-table RLS as
--       defense-in-depth, so a direct REST hit on income/expenses is gated too.
-- Re-assert anon revoke on every rebuilt view (regression guard).
--
-- OPEN DECISIONS for review (chose the conservative option; change if desired):
--   D1: income/expenses WRITES are master_admin-only here. Effect: the Events
--       hub "add income / add expense" controls become master-only. If you want
--       operations staff to log event expenses pre-audit, widen the write
--       policy to can_use_events_module().
--   D2: ticketing / sponsorships / third_party_donations WRITES are left to
--       can_use_events_module() (operational entry happens in the hub); their
--       READS are audited-gated for staff. Flip to master-only if you'd rather.
-- =============================================================================
begin;

-- 1. Audit flag on events -----------------------------------------------------
alter table public.events
  add column if not exists audited     boolean not null default false,
  add column if not exists audited_at  timestamptz,
  add column if not exists audited_by  uuid references public.profiles(id);

-- Only master_admin may set/clear audited. Guard via a trigger so it survives
-- whatever the events module write policy allows.
create or replace function public.guard_event_audit()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if (coalesce(new.audited,false) is distinct from coalesce(old.audited,false))
     and not public.is_master_admin() then
    raise exception 'Only master_admin may change an event audit status';
  end if;
  if new.audited and not coalesce(old.audited,false) then
    new.audited_at := now(); new.audited_by := auth.uid();
  elsif not new.audited then
    new.audited_at := null; new.audited_by := null;
  end if;
  return new;
end $$;

drop trigger if exists trg_guard_event_audit on public.events;
create trigger trg_guard_event_audit
  before update on public.events
  for each row execute function public.guard_event_audit();

-- 2. Visibility helper: may the current user see THIS event's financials? ------
create or replace function public.can_see_event_financials(p_event_id uuid)
returns boolean
language sql stable security definer set search_path = public
as $$
  select public.is_master_admin()
      or (p_event_id is not null
          and exists (select 1 from public.events e
                       where e.id = p_event_id and e.audited));
$$;
grant execute on function public.can_see_event_financials(uuid) to authenticated;

-- 3. Base-table RLS (defense-in-depth) ----------------------------------------
-- income: company-level rows (event_id is null) -> master only; event rows ->
-- audited gate. Writes master-only (D1).
drop policy if exists "Admins can manage income" on public.income;
create policy "Income read gated" on public.income for select
  using (public.can_see_event_financials(event_id));
create policy "Income write master" on public.income for all
  using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Admins can manage expenses" on public.expenses;
create policy "Expenses read gated" on public.expenses for select
  using (public.can_see_event_financials(event_id));
create policy "Expenses write master" on public.expenses for all
  using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Admins can manage mileage" on public.mileage;
create policy "Mileage master only" on public.mileage for all
  using (public.is_master_admin()) with check (public.is_master_admin());

-- ticketing / sponsorships / third_party_donations: read audited-gated,
-- write via events module (D2).
drop policy if exists "Admins can manage ticketing" on public.ticketing;
create policy "Ticketing read gated" on public.ticketing for select
  using (public.can_see_event_financials(event_id));
create policy "Ticketing write events" on public.ticketing for all
  using (public.can_use_events_module()) with check (public.can_use_events_module());

drop policy if exists "Admins can manage sponsorships" on public.sponsorships;
create policy "Sponsorships read gated" on public.sponsorships for select
  using (public.can_see_event_financials(event_id));
create policy "Sponsorships write events" on public.sponsorships for all
  using (public.can_use_events_module()) with check (public.can_use_events_module());

drop policy if exists "Admins can manage third-party donations" on public.third_party_donations;
create policy "Donations read gated" on public.third_party_donations for select
  using (public.can_see_event_financials(event_id));
create policy "Donations write events" on public.third_party_donations for all
  using (public.can_use_events_module()) with check (public.can_use_events_module());

-- 4. Rebuild financial views: CASE-gate the money columns + security_invoker.
--    Definitions mirror 022; only the financial columns are wrapped in a gate
--    of (is_master_admin() OR e.audited). Non-financial columns stay visible.
create or replace view public.v_event_summary
with (security_invoker = true) as
select
  e.id as event_id, e.slug, e.name, e.event_date, e.series, e.status, e.venue_id,
  case when g then coalesce(rev.revenue,0) end                              as revenue,
  case when g then coalesce(exp.expenses,0) end                            as expenses,
  case when g then coalesce(rev.revenue,0)-coalesce(exp.expenses,0) end    as net,
  coalesce(spn.sponsor_count,0) as sponsor_count,
  case when g then coalesce(spn.sponsor_cash,0) end                        as sponsor_cash,
  coalesce(tkt.tickets_sold,0)  as tickets_sold,
  case when g then coalesce(tkt.ticket_revenue,0) end                      as ticket_revenue,
  e.total_attendance,
  case when g then coalesce(dn.third_party_total,0) end                    as third_party_donations,
  case when g then coalesce(spn.sponsor_in_kind,0) end                     as sponsor_in_kind
from public.events e
cross join lateral (select (public.is_master_admin() or e.audited) as g) gate
left join lateral (select sum(amount) revenue  from public.income   where event_id=e.id and deleted_at is null) rev on true
left join lateral (select sum(amount) expenses from public.expenses where event_id=e.id and deleted_at is null) exp on true
left join lateral (select count(*) sponsor_count, sum(cash_amount) sponsor_cash, sum(in_kind_value) sponsor_in_kind
                   from public.sponsorships where event_id=e.id and status <> 'cancelled') spn on true
left join lateral (select sum(coalesce(quantity,1)) tickets_sold, sum(amount_paid) ticket_revenue
                   from public.ticketing where event_id=e.id) tkt on true
left join lateral (select sum(amount) third_party_total from public.third_party_donations where event_id=e.id) dn on true
where e.deleted_at is null;
revoke select on public.v_event_summary from anon;

-- v_kpi_event_financials, v_kpi_parties, v_kpi_dance_infusion: rebuild exactly
-- as 022 but reading from the now-gated v_event_summary (so they inherit the
-- per-event gate). v_kpi_dashboard stays master-only (strategy rollup).
drop view if exists public.v_kpi_parties;
drop view if exists public.v_kpi_dance_infusion;
drop view if exists public.v_kpi_event_financials;

create view public.v_kpi_event_financials with (security_invoker = true) as
select s.event_id, s.name, s.series, s.event_date, e.capacity, s.total_attendance,
  s.tickets_sold, s.ticket_revenue, s.revenue as other_income, s.expenses as total_expenses,
  s.third_party_donations as donations, s.sponsor_cash, s.sponsor_in_kind,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash)              as gross_revenue,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash) - s.expenses as net_pl,
  (s.ticket_revenue + s.revenue + s.third_party_donations + s.sponsor_cash + s.sponsor_in_kind) as total_raised
from public.v_event_summary s join public.events e on e.id = s.event_id;
revoke select on public.v_kpi_event_financials from anon;

create view public.v_kpi_parties with (security_invoker = true) as
select event_id, name, event_date, capacity, tickets_sold,
  case when capacity > 0 then round(tickets_sold::numeric/capacity*100,1) end as sell_through_pct,
  net_pl
from public.v_kpi_event_financials where series = 'Come With Parties';
revoke select on public.v_kpi_parties from anon;

create view public.v_kpi_dance_infusion with (security_invoker = true) as
select event_id, name, event_date, total_attendance, net_pl, total_raised,
  case when total_raised > 0 then round(total_expenses/nullif(total_raised,0),2) end as cost_to_raise_per_dollar
from public.v_kpi_event_financials where series = 'Dance Infusion';
revoke select on public.v_kpi_dance_infusion from anon;

-- v_kpi_dashboard (strategy rollup) — keep master-only. If it is not already
-- security_invoker, leave its definition and just ensure anon is revoked.
revoke select on public.v_kpi_dashboard from anon;

commit;

-- =============================================================================
-- POST-APPLY VERIFICATION:
--   * anon REST GET on all five v_* financial views -> 401 (regression guard).
--   * master (berky): sees full financials on every event.
--   * staff (events access, event.audited=false): financial columns are NULL /
--     event income+expenses return 0 rows. Flip audited=true -> staff now see
--     that event's final numbers.
--   * income/expenses with event_id IS NULL: invisible to staff in all cases.
-- =============================================================================
