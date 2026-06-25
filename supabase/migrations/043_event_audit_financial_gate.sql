-- =============================================================================
-- 043_event_audit_financial_gate.sql  (REWRITTEN 2026-06-25)
-- Two-flag, event-scoped financial publish gate + master/sub financial split.
--
-- ┌──────────────────────────────────────────────────────────────────────────┐
-- │  NOT YET APPLIED TO PROD. Review + anon-revoke verification required       │
-- │  (the 016/017->019 regression class lives exactly here). Apply 042 first.  │
-- └──────────────────────────────────────────────────────────────────────────┘
--
-- MODEL (per Keith, 2026-06-25):
--   * events.audited (master-only): "I've checked these numbers." Informational
--     — it does NOT itself reveal anything. Drives the UI warning severity.
--   * events.financials_released (master-only): the switch that actually reveals
--     THIS event's financials to staff. Always selectable; the dashboard pops a
--     confirm every time, and a LOUD red warning when the event is not audited.
--   * Staff see an event's money only when financials_released = true. Company-
--     level rows (event_id IS NULL = overhead / general income) are master-only
--     PERMANENTLY (no flag releases them).
--
-- Why the gate must live in BOTH layers (from the audit): the v_* money views
-- were created WITHOUT security_invoker (015/022/051) so they run as owner and
-- bypass base-table RLS. Two-part fix:
--   (A) base-table RLS on income/expenses/mileage/ticketing/sponsorships/
--       donations (defense-in-depth: a direct REST hit is gated), and
--   (B) gate the money COLUMNS inside v_event_summary with a CASE, and set every
--       money view security_invoker = true so they honour that base RLS. The
--       downstream views (v_kpi_event_financials -> parties/dance_infusion ->
--       v_kpi_computed -> v_kpi_dashboard) inherit the gate because they read
--       through v_event_summary; we DON'T redefine them (avoids the 051 cascade)
--       — just flip security_invoker and re-assert the anon revoke.
--
-- Writes (review-time decisions, conservative defaults):
--   D1 income / expenses / mileage WRITES = master_admin only (company P&L).
--   D2 ticketing / sponsorships / donations WRITES = can_use_events_module()
--      (operational entry happens in the Events hub). READS audited/release-gated.
-- =============================================================================
begin;

-- 1. Audit + release flags on events ------------------------------------------
alter table public.events
  add column if not exists audited                 boolean not null default false,
  add column if not exists audited_at              timestamptz,
  add column if not exists audited_by              uuid references public.profiles(id),
  add column if not exists financials_released     boolean not null default false,
  add column if not exists financials_released_at  timestamptz,
  add column if not exists financials_released_by  uuid references public.profiles(id);

-- Only master_admin may flip either flag; stamp who/when. Survives whatever the
-- events module write policy allows (guard runs security definer on every write).
create or replace function public.guard_event_finance_flags()
returns trigger language plpgsql security definer set search_path = public as $$
declare
  old_aud boolean := coalesce(case when tg_op = 'UPDATE' then old.audited end, false);
  old_rel boolean := coalesce(case when tg_op = 'UPDATE' then old.financials_released end, false);
begin
  if (coalesce(new.audited, false) is distinct from old_aud
      or coalesce(new.financials_released, false) is distinct from old_rel)
     and not public.is_master_admin() then
    raise exception 'Only master_admin may change event audit / financial-release status';
  end if;
  if coalesce(new.audited, false) and not old_aud then
    new.audited_at := now(); new.audited_by := auth.uid();
  elsif not coalesce(new.audited, false) then
    new.audited_at := null; new.audited_by := null;
  end if;
  if coalesce(new.financials_released, false) and not old_rel then
    new.financials_released_at := now(); new.financials_released_by := auth.uid();
  elsif not coalesce(new.financials_released, false) then
    new.financials_released_at := null; new.financials_released_by := null;
  end if;
  return new;
end $$;
drop trigger if exists trg_guard_event_finance_flags on public.events;
create trigger trg_guard_event_finance_flags
  before insert or update on public.events
  for each row execute function public.guard_event_finance_flags();

-- 2. Visibility helper: may the current user see THIS event's financials? ------
create or replace function public.can_see_event_financials(p_event_id uuid)
returns boolean language sql stable security definer set search_path = public
as $$
  select public.is_master_admin()
      or (p_event_id is not null
          and exists (select 1 from public.events e
                       where e.id = p_event_id and e.financials_released));
$$;
grant execute on function public.can_see_event_financials(uuid) to authenticated;

-- 3. Base-table RLS (defense-in-depth) ----------------------------------------
-- income / expenses / mileage: reads release-gated (event_id null => master);
-- writes master-only (D1).
drop policy if exists "Admins can manage income" on public.income;
create policy "Income read gated"  on public.income for select using (public.can_see_event_financials(event_id));
create policy "Income write master" on public.income for all using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Admins can manage expenses" on public.expenses;
create policy "Expenses read gated"  on public.expenses for select using (public.can_see_event_financials(event_id));
create policy "Expenses write master" on public.expenses for all using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Admins can manage mileage" on public.mileage;
create policy "Mileage master only" on public.mileage for all using (public.is_master_admin()) with check (public.is_master_admin());

-- ticketing / sponsorships / donations: reads release-gated; writes via Events hub (D2).
drop policy if exists "Admins can manage ticketing" on public.ticketing;
create policy "Ticketing read gated"  on public.ticketing for select using (public.can_see_event_financials(event_id));
create policy "Ticketing write events" on public.ticketing for all using (public.can_use_events_module()) with check (public.can_use_events_module());

drop policy if exists "Admins can manage sponsorships" on public.sponsorships;
create policy "Sponsorships read gated"  on public.sponsorships for select using (public.can_see_event_financials(event_id));
create policy "Sponsorships write events" on public.sponsorships for all using (public.can_use_events_module()) with check (public.can_use_events_module());

drop policy if exists "Admins can manage third-party donations" on public.third_party_donations;
create policy "Donations read gated"  on public.third_party_donations for select using (public.can_see_event_financials(event_id));
create policy "Donations write events" on public.third_party_donations for all using (public.can_use_events_module()) with check (public.can_use_events_module());

-- 4. v_event_summary: gate the money columns + security_invoker. Definition is
--    the live (051-era) one, with money columns wrapped in a CASE on the gate.
--    Non-money columns (counts, attendance, identity) stay visible.
create or replace view public.v_event_summary with (security_invoker = true) as
select
  e.id as event_id, e.slug, e.name, e.event_date, e.series, e.status, e.venue_id,
  case when g.ok then coalesce(rev.revenue, 0::numeric) end                                    as revenue,
  case when g.ok then coalesce(exp.expenses, 0::numeric) end                                   as expenses,
  case when g.ok then coalesce(rev.revenue, 0::numeric) - coalesce(exp.expenses, 0::numeric) end as net,
  coalesce(spn.sponsor_count, 0::bigint) as sponsor_count,
  case when g.ok then coalesce(spn.sponsor_cash, 0::numeric) end                               as sponsor_cash,
  coalesce(tkt.tickets_sold, 0::bigint)  as tickets_sold,
  case when g.ok then coalesce(tkt.ticket_revenue, 0::numeric) end                             as ticket_revenue,
  e.total_attendance,
  case when g.ok then coalesce(dn.third_party_total, 0::numeric) end                           as third_party_donations,
  case when g.ok then coalesce(spn.sponsor_in_kind, 0::numeric) end                            as sponsor_in_kind
from public.events e
cross join lateral (select (public.is_master_admin() or e.financials_released) as ok) g
left join lateral (select sum(income.amount) as revenue   from public.income   where income.event_id = e.id and income.deleted_at is null) rev on true
left join lateral (select sum(expenses.amount) as expenses from public.expenses where expenses.event_id = e.id and expenses.deleted_at is null) exp on true
left join lateral (select count(*) as sponsor_count, sum(sponsorships.cash_amount) as sponsor_cash, sum(sponsorships.in_kind_value) as sponsor_in_kind
                   from public.sponsorships where sponsorships.event_id = e.id and sponsorships.status <> 'cancelled'::text) spn on true
left join lateral (select sum(coalesce(ticketing.quantity, 1)) as tickets_sold, sum(ticketing.amount_paid) as ticket_revenue
                   from public.ticketing where ticketing.event_id = e.id) tkt on true
left join lateral (select sum(third_party_donations.amount) as third_party_total from public.third_party_donations where third_party_donations.event_id = e.id) dn on true
where e.deleted_at is null;

-- 5. Flip security_invoker on the downstream money views (inherit the gate via
--    v_event_summary; no redefinition, so 051's v_kpi_computed/dashboard stay
--    intact). Re-assert the anon revoke on every one (regression guard).
alter view public.v_kpi_event_financials set (security_invoker = true);
alter view public.v_kpi_parties          set (security_invoker = true);
alter view public.v_kpi_dance_infusion   set (security_invoker = true);
alter view public.v_kpi_computed         set (security_invoker = true);
alter view public.v_kpi_dashboard        set (security_invoker = true);

revoke select on public.v_event_summary        from anon;
revoke select on public.v_kpi_event_financials from anon;
revoke select on public.v_kpi_parties          from anon;
revoke select on public.v_kpi_dance_infusion   from anon;
revoke select on public.v_kpi_computed         from anon;
revoke select on public.v_kpi_dashboard        from anon;

commit;

-- =============================================================================
-- POST-APPLY VERIFICATION:
--   * anon REST GET on all six v_* views -> 401 (regression guard).
--   * master (berky): full financials on every event; can flip both flags.
--   * staff (events access, financials_released=false): money columns NULL and
--     income/expenses return 0 rows. Flip financials_released=true -> that
--     event's numbers appear. event_id IS NULL rows never appear for staff.
--   * non-master attempting to set audited/financials_released -> exception.
-- FOLLOW-UP (not in this migration; tracked in ROADMAP "GATED BLOCKER"):
--   v_budget_variance, v_data_points, mv_event_data_points still need locking
--   before the financial-view blocker is fully closed (the MV can't use RLS —
--   revoke from authenticated + expose via a gated view or service-role only).
-- =============================================================================
