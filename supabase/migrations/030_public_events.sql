-- =============================================================================
-- 030_public_events.sql
-- Public events on comewith.org: a data-driven, anon-readable list of UPCOMING
-- events with an external ticket/RSVP link, fed from the admin dashboard.
--
-- TWO clearly separated concerns:
--
--   A) SECURITY RE-LOCK — close a pre-existing anon exposure on the events TABLE.
--      013's blanket `grant all on all tables to anon`, combined with the
--      "Public can read non-cancelled events" RLS policy, let anon read EVERY
--      column of every non-cancelled event (bar_minimum, notes, total_attendance,
--      capacity, type, …). 019 only re-revoked the financial VIEWS; it never
--      touched the events TABLE. Nothing public depends on direct table reads —
--      the DI hub (get-event-hub) uses the service role and bypasses RLS — so we
--      revoke anon from the table and drop the public-read policy. Admins
--      (authenticated + is_admin()) are unaffected; the `authenticated` table
--      grant is left intact so admin reads still reach policy evaluation.
--
--   B) FEATURE — add is_public + ticket_label, then expose ONLY safe fields
--      through a dedicated view and grant anon SELECT on THAT VIEW ONLY.
--
-- Reversible: see 030_public_events.down.sql. DOWN drops the view + new columns
-- but intentionally KEEPS the security re-lock — re-opening the hole would be a
-- regression.
-- =============================================================================
begin;

-- ── A) Security re-lock: the events TABLE is admin-only again ─────────────────
revoke all on public.events from anon;
drop policy if exists "Public can read non-cancelled events" on public.events;

-- ── B) New fields (ticket_url already exists from 007_events.sql) ─────────────
alter table public.events
  add column if not exists is_public    boolean not null default false,
  add column if not exists ticket_label text;

comment on column public.events.is_public is
  'When true, the event surfaces on the public site via v_public_events (future-dated, non-cancelled only). Defaults false so nothing leaks by accident.';
comment on column public.events.ticket_label is
  'Button text for the public ticket/RSVP link (e.g. "Get tickets", "RSVP"). UI falls back to a default when null.';

-- ── B) Dedicated public view: ONLY safe fields, future + live events ─────────
-- venue NAME only — never address / contact / capacity. No financials, notes,
-- attendance, participants, or internal fields.
create or replace view public.v_public_events as
  select
    e.name,
    e.event_date,
    v.name        as venue_name,
    e.ticket_url,
    e.ticket_label
  from public.events e
  left join public.venues v on v.id = e.venue_id
  where e.is_public  = true
    and e.event_date >= current_date
    and e.deleted_at is null
    and e.status    <> 'cancelled';   -- defensive: never show a cancelled event

comment on view public.v_public_events is
  'Anon-readable public events feed for comewith.org. Exposes ONLY name / event_date / venue_name / ticket_url / ticket_label for is_public, future-dated, non-cancelled, non-deleted events. The events table itself is anon-revoked.';

-- ── B) Least privilege: strip the auto-grant 013 default privileges apply to
--       new views, then grant SELECT only. This is the one anon-readable object.
revoke all    on public.v_public_events from anon, authenticated;
grant  select on public.v_public_events to   anon, authenticated;

commit;
