-- =============================================================================
-- 042_module_rls_gate.sql
-- Hard RLS enforcement of the staff module gate (the "RLS-enforced per module"
-- choice). NON-FINANCIAL TABLES ONLY.
--
-- ┌──────────────────────────────────────────────────────────────────────────┐
-- │  NOT YET APPLIED TO PROD.  Review required before applying.                │
-- │  Reason it is staged, not auto-applied:                                    │
-- │   The Events module is SIGNED OFF and its hub reads ~15 other tables       │
-- │   (venues, sponsors, artists, equipment, tasks, contracts, files, …). A    │
-- │   naive "one SELECT policy per module" would break the Events hub for the  │
-- │   live sub_admin (liz@comewith.org) the instant it lands — a prod          │
-- │   regression that git cannot roll back. This migration solves that with an │
-- │   events-dependency carve: any table the hub reads is also readable/        │
-- │   writable by whoever can access the Events module. Apply only after a     │
-- │   read-through + a staff-account smoke test.                               │
-- └──────────────────────────────────────────────────────────────────────────┘
--
-- Pattern, per gated table T owned by module K:
--   * DROP the old blanket "Admins can manage T" (== is_admin) policy.
--   * KEEP every public/anon/customer policy untouched (inquiry insert, public
--     event/equipment read, agreement-by-token, customer read-own, …).
--   * ADD a module-gated policy. user_can_access_module() already returns TRUE
--     for master_admin, so master keeps full access automatically and the
--     master/sub split falls out for free (sub is now gated, master is not).
--   * For Events-hub dependency tables, OR-in user_can_access_module('events').
--
-- FINANCIAL tables (income, expenses, mileage, ticketing, sponsorships,
-- third_party_donations, raffle_prizes) and the v_* financial views are NOT
-- touched here — they are handled in 043 under the event audit gate.
-- =============================================================================
begin;

-- Helper: may the current user act within the (signed-off) Events module?
-- Used to carve read/write on every table the Events hub touches so that
-- gating those tables for their *own* module does not break the hub.
create or replace function public.can_use_events_module()
returns boolean
language sql stable security definer set search_path = public
as $$ select public.user_can_access_module('events') $$;
grant execute on function public.can_use_events_module() to authenticated;

-- ── Sales ────────────────────────────────────────────────────────────────────
-- inquiries (own module only; not a hub dependency). Keep anon insert.
drop policy if exists "Admins can read all inquiries"   on public.inquiries;
drop policy if exists "Admins can update inquiries"      on public.inquiries;
drop policy if exists "Master admin can delete inquiries" on public.inquiries;
create policy "Inquiries module access" on public.inquiries for select
  using (public.user_can_access_module('inquiries'));
create policy "Inquiries module update" on public.inquiries for update
  using (public.user_can_access_module('inquiries'));
create policy "Inquiries module delete" on public.inquiries for delete
  using (public.is_master_admin());

-- agreements (hub reads contracts/agreements). Keep customer read-own.
drop policy if exists "Admins can manage all agreements" on public.agreements;
create policy "Agreements module access" on public.agreements for all
  using (public.user_can_access_module('agreements') or public.can_use_events_module())
  with check (public.user_can_access_module('agreements') or public.can_use_events_module());

-- clients (reached via agreements in the hub). Keep customer read-own.
drop policy if exists "Admins can manage clients" on public.clients;
create policy "Clients module access" on public.clients for all
  using (public.user_can_access_module('clients') or public.can_use_events_module())
  with check (public.user_can_access_module('clients') or public.can_use_events_module());

-- ── Operations ───────────────────────────────────────────────────────────────
-- events itself. Keep "Public can read non-cancelled events".
drop policy if exists "Admins can manage events" on public.events;
create policy "Events module access" on public.events for all
  using (public.user_can_access_module('events'))
  with check (public.user_can_access_module('events'));

-- venues (hub dependency). Keep any public read if present.
drop policy if exists "Admins can manage venues" on public.venues;
create policy "Venues module access" on public.venues for all
  using (public.user_can_access_module('venues') or public.can_use_events_module())
  with check (public.user_can_access_module('venues') or public.can_use_events_module());

-- equipment (hub uses equipment_usage -> equipment_inventory). Keep public read.
drop policy if exists "Admins can manage equipment" on public.equipment_inventory;
create policy "Equipment module access" on public.equipment_inventory for all
  using (public.user_can_access_module('equipment') or public.can_use_events_module())
  with check (public.user_can_access_module('equipment') or public.can_use_events_module());
drop policy if exists "Admins can manage equipment usage" on public.equipment_usage;
create policy "Equipment usage module access" on public.equipment_usage for all
  using (public.user_can_access_module('equipment') or public.can_use_events_module())
  with check (public.user_can_access_module('equipment') or public.can_use_events_module());

-- templates  [VERIFY: confirm the table is public.task_templates and the exact
-- existing policy name before applying — migration 026 created it.]
drop policy if exists "Admins can manage task templates" on public.task_templates;
create policy "Templates module access" on public.task_templates for all
  using (public.user_can_access_module('templates') or public.can_use_events_module())
  with check (public.user_can_access_module('templates') or public.can_use_events_module());

-- ── Partners ────────────────────────────────────────────────────────────────
-- sponsors (hub dependency; non-financial — money lives on sponsorships=043).
drop policy if exists "Admins can manage sponsors" on public.sponsors;
create policy "Sponsors module access" on public.sponsors for all
  using (public.user_can_access_module('sponsors') or public.can_use_events_module())
  with check (public.user_can_access_module('sponsors') or public.can_use_events_module());

-- artists + bookings (hub dependency). artist_notes already has its own
-- master/sub split (008) — leave it.
drop policy if exists "Admins can manage artists" on public.artists;
create policy "Artists module access" on public.artists for all
  using (public.user_can_access_module('artists') or public.can_use_events_module())
  with check (public.user_can_access_module('artists') or public.can_use_events_module());
drop policy if exists "Admins can manage artist bookings" on public.artist_bookings;
create policy "Artist bookings module access" on public.artist_bookings for all
  using (public.user_can_access_module('artists') or public.can_use_events_module())
  with check (public.user_can_access_module('artists') or public.can_use_events_module());

-- ── Audience ─────────────────────────────────────────────────────────────────
-- guests (ticketing -> guests in the hub). Edge functions use service_role.
drop policy if exists "Admins can manage guests" on public.guests;
create policy "Guests module access" on public.guests for all
  using (public.user_can_access_module('guests') or public.can_use_events_module())
  with check (public.user_can_access_module('guests') or public.can_use_events_module());

-- subscribers + segments + campaigns (own module; not hub dependencies).
-- [VERIFY existing policy names from 009_mailing_list.sql before applying.]
drop policy if exists "Admins can manage subscribers" on public.subscribers;
create policy "Subscribers module access" on public.subscribers for all
  using (public.user_can_access_module('subscribers'))
  with check (public.user_can_access_module('subscribers'));
drop policy if exists "Admins can manage mailing campaigns" on public.mailing_campaigns;
create policy "Campaigns module access" on public.mailing_campaigns for all
  using (public.user_can_access_module('campaigns'))
  with check (public.user_can_access_module('campaigns'));

-- ── Insights ─────────────────────────────────────────────────────────────────
-- notes == feedback_log.  [VERIFY policy name from 016_feedback_log.sql.]
drop policy if exists "Admins can manage feedback_log" on public.feedback_log;
create policy "Notes module access" on public.feedback_log for all
  using (public.user_can_access_module('notes'))
  with check (public.user_can_access_module('notes'));

commit;

-- =============================================================================
-- POST-APPLY SMOKE TEST (do this with a real 'operations' staff account, e.g.
-- a throwaway, BEFORE trusting it):
--   * master_admin (berky): can still read/write every table above.
--   * staff with events access: can open an Events hub and see venues/sponsors/
--     artists/equipment/tasks/contracts/files (the carve works).
--   * staff WITHOUT a module: select returns 0 rows / 401, not an error.
--   * anon: inquiry insert still works; public event/equipment read still works.
-- ROLLBACK: re-create each "Admins can manage X" policy as `for all using
--   (public.is_admin())` and drop the "<Module> module access" policies.
-- =============================================================================
