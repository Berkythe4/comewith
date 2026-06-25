-- =============================================================================
-- 042_module_rls_gate.sql   (REWRITTEN 2026-06-25 for the post-047 actor model)
-- Hard RLS enforcement of the staff module gate. NON-FINANCIAL TABLES ONLY.
--
-- This supersedes the original 2026-06-23 draft, which referenced clients /
-- sponsors / artists / artist_bookings — all DROPPED in 047. Those are now the
-- actor model (actors / actor_roles / event_participants), gated below.
--
-- Pattern per gated table T owned by module K:
--   * DROP the old blanket "Admins can manage T" (is_admin) policy.
--   * KEEP every public / anon / customer policy untouched (inquiry insert,
--     public event/equipment read, agreement-by-token, customer/actor read-own).
--   * ADD a module-gated policy. user_can_access_module() returns TRUE for
--     master_admin, so master keeps full access automatically; the master/sub
--     split falls out for free (sub is gated only to signed-off modules in scope).
--   * For tables the Events hub reads, OR-in can_use_events_module() so gating a
--     table for its own module never breaks the (signed-off) Events hub.
--
-- FINANCIAL tables (income, expenses, mileage, ticketing, sponsorships,
-- third_party_donations) and the v_* money views are NOT touched here — they are
-- handled in 043 under the event audit / financial-release gate.
-- =============================================================================
begin;

-- Helper: may the current user act within the (signed-off) Events module?
create or replace function public.can_use_events_module()
returns boolean language sql stable security definer set search_path = public
as $$ select public.user_can_access_module('events') $$;
grant execute on function public.can_use_events_module() to authenticated;

-- Helper: the contact roster (actors / roles / participation) is read by the
-- Actors, Clients, Sponsors and Artists tabs AND the Events hub. Anyone with any
-- of those modules (or the hub) may use it. Master is always TRUE.
create or replace function public.can_see_people()
returns boolean language sql stable security definer set search_path = public
as $$
  select public.user_can_access_module('actors')
      or public.user_can_access_module('clients')
      or public.user_can_access_module('sponsors')
      or public.user_can_access_module('artists')
      or public.can_use_events_module()
$$;
grant execute on function public.can_see_people() to authenticated;

-- ── Sales ────────────────────────────────────────────────────────────────────
-- inquiries (own module; keep "Anyone can submit an inquiry" INSERT).
drop policy if exists "Admins can read all inquiries"    on public.inquiries;
drop policy if exists "Admins can update inquiries"       on public.inquiries;
drop policy if exists "Master admin can delete inquiries" on public.inquiries;
create policy "Inquiries module select" on public.inquiries for select
  using (public.user_can_access_module('inquiries'));
create policy "Inquiries module update" on public.inquiries for update
  using (public.user_can_access_module('inquiries'))
  with check (public.user_can_access_module('inquiries'));
create policy "Inquiries module delete" on public.inquiries for delete
  using (public.is_master_admin());

-- agreements (Events hub reads them; keep "Customers can read own agreements").
drop policy if exists "Admins can manage all agreements" on public.agreements;
create policy "Agreements module access" on public.agreements for all
  using (public.user_can_access_module('agreements') or public.can_use_events_module())
  with check (public.user_can_access_module('agreements') or public.can_use_events_module());

-- ── People / actor model (replaces the retired clients/sponsors/artists) ──────
drop policy if exists "Admins can manage actors" on public.actors;
create policy "Actors module access" on public.actors for all
  using (public.can_see_people()) with check (public.can_see_people());

drop policy if exists "Admins can manage actor roles" on public.actor_roles;
create policy "Actor roles module access" on public.actor_roles for all
  using (public.can_see_people()) with check (public.can_see_people());

drop policy if exists "Admins can manage event participants" on public.event_participants;
create policy "Event participants module access" on public.event_participants for all
  using (public.user_can_access_module('actors') or public.can_use_events_module())
  with check (public.user_can_access_module('actors') or public.can_use_events_module());

-- ── Operations ───────────────────────────────────────────────────────────────
drop policy if exists "Admins can manage events" on public.events;
create policy "Events module access" on public.events for all
  using (public.user_can_access_module('events'))
  with check (public.user_can_access_module('events'));

drop policy if exists "Admins can manage venues" on public.venues;
create policy "Venues module access" on public.venues for all
  using (public.user_can_access_module('venues') or public.can_use_events_module())
  with check (public.user_can_access_module('venues') or public.can_use_events_module());

-- equipment (hub uses equipment_usage -> equipment_inventory; keep public read).
drop policy if exists "Admins can manage equipment" on public.equipment_inventory;
create policy "Equipment module access" on public.equipment_inventory for all
  using (public.user_can_access_module('equipment') or public.can_use_events_module())
  with check (public.user_can_access_module('equipment') or public.can_use_events_module());
drop policy if exists "Admins can manage equipment usage" on public.equipment_usage;
create policy "Equipment usage module access" on public.equipment_usage for all
  using (public.user_can_access_module('equipment') or public.can_use_events_module())
  with check (public.user_can_access_module('equipment') or public.can_use_events_module());

-- templates (task_templates; hub generates day-of tasks from them).
drop policy if exists "Admins can manage task templates" on public.task_templates;
create policy "Templates module access" on public.task_templates for all
  using (public.user_can_access_module('templates') or public.can_use_events_module())
  with check (public.user_can_access_module('templates') or public.can_use_events_module());

-- contracts (Events-hub only; keep "Actors can read own contracts").
drop policy if exists "Admins can manage contracts" on public.contracts;
create policy "Contracts module access" on public.contracts for all
  using (public.can_use_events_module()) with check (public.can_use_events_module());

-- files (Events hub + agreement attachments; service-role Edge fns bypass RLS).
drop policy if exists "Admins can manage files" on public.files;
create policy "Files module access" on public.files for all
  using (public.can_use_events_module() or public.user_can_access_module('agreements'))
  with check (public.can_use_events_module() or public.user_can_access_module('agreements'));

-- document_types (lookup for the Files buckets).
drop policy if exists "Admins manage document types" on public.document_types;
create policy "Document types module access" on public.document_types for all
  using (public.can_use_events_module()) with check (public.can_use_events_module());

-- ── Audience ─────────────────────────────────────────────────────────────────
-- guests (ticketing -> guests in the hub; Edge fns use service_role).
drop policy if exists "Admins can manage guests" on public.guests;
create policy "Guests module access" on public.guests for all
  using (public.user_can_access_module('guests') or public.can_use_events_module())
  with check (public.user_can_access_module('guests') or public.can_use_events_module());

-- subscribers (own module; keep "Anyone can subscribe" INSERT).
drop policy if exists "Admins can manage subscribers" on public.subscribers;
create policy "Subscribers module access" on public.subscribers for all
  using (public.user_can_access_module('subscribers'))
  with check (public.user_can_access_module('subscribers'));

-- mailing_campaigns (own module).
drop policy if exists "Admins can manage mailing campaigns" on public.mailing_campaigns;
create policy "Campaigns module access" on public.mailing_campaigns for all
  using (public.user_can_access_module('campaigns'))
  with check (public.user_can_access_module('campaigns'));

-- ── Insights ─────────────────────────────────────────────────────────────────
-- notes == feedback_log.
drop policy if exists "Admins can manage feedback_log" on public.feedback_log;
create policy "Notes module access" on public.feedback_log for all
  using (public.user_can_access_module('notes'))
  with check (public.user_can_access_module('notes'));

commit;

-- =============================================================================
-- POST-APPLY SMOKE TEST (run with a throwaway 'operations' staff account BEFORE
-- trusting it):
--   * master_admin (berky): still reads/writes every table above.
--   * staff w/ Events access: opens an Events hub and sees venues / people /
--     equipment / contracts / files (the can_use_events_module carve works).
--   * staff WITHOUT a signed-off module: select returns 0 rows / 401, not error.
--   * anon: inquiry insert + public event/equipment read still work.
-- NOTE: social_posts/social_post_notes already have per-module RLS (044) — not
--   touched here. subscriber_segments inherits its existing policy (low-risk).
-- ROLLBACK: drop each "<X> module access" policy and recreate the original
--   "Admins can manage X" as `for all using (public.is_admin())`.
-- =============================================================================
