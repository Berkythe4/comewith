-- =============================================================================
-- 057_fix_financial_read_leak.sql
-- BUG in 043: ticketing / sponsorships / third_party_donations got a WRITE policy
-- `for all using can_use_events_module()`. Postgres `for all` ALSO covers SELECT,
-- so the permissive OR meant any events-module staff could read these revenue
-- tables via direct REST — bypassing the "read gated" policy. (income/expenses
-- were safe because their write policy uses is_master_admin().)
--
-- Fix: make the three write policies master-only, matching income/expenses. Reads
-- stay governed by the "<X> read gated" = can_see_event_financials() policy, so a
-- staffer sees a row only for an event whose financials are RELEASED.
-- =============================================================================
begin;
drop policy if exists "Ticketing write events" on public.ticketing;
create policy "Ticketing write master" on public.ticketing for all
  using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Sponsorships write events" on public.sponsorships;
create policy "Sponsorships write master" on public.sponsorships for all
  using (public.is_master_admin()) with check (public.is_master_admin());

drop policy if exists "Donations write events" on public.third_party_donations;
create policy "Donations write master" on public.third_party_donations for all
  using (public.is_master_admin()) with check (public.is_master_admin());
commit;
-- POST: a non-master staffer reading ticketing / sponsorships / third_party_donations
-- for an unreleased event returns 0 rows. Master + released events unchanged.
