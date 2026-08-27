-- =============================================================================
-- 204_artist_gigs_public_only.sql
-- A gig shows on a public artist profile only if the EVENT is public.
--
-- v_artist_gigs (065) read `is_public = true OR status = 'completed'`. The
-- second half was the leak: every completed event listed every participant by
-- name on artist.html, whether or not the event was ever announced publicly.
-- Private bookings, corporate gigs and anything Keith deliberately left
-- unpublished were all on the site the moment they were marked complete.
--
-- `is_public` is the one deliberate decision about whether an event faces the
-- public, so it is the only gate here now. A gig that should appear and doesn't
-- is fixed by toggling `is_public` on the event — not by widening this view.
--
-- The view is NOT security_invoker, so it runs as owner and RLS on `events`
-- never applies to it: this WHERE clause is the whole gate. Columns are
-- unchanged, so `create or replace` keeps the grants from 065 (anon reads it,
-- by design — it is a public profile page).
--
-- Consumers: artist.html only (grep confirmed). No dashboard surface reads it.
-- =============================================================================
begin;

create or replace view public.v_artist_gigs as
  select ep.actor_id, e.id as event_id, e.name as event_name, e.event_date,
         v.name as venue_name, ep.role
  from public.event_participants ep
  join public.events e on e.id = ep.event_id and e.deleted_at is null
  left join public.venues v on v.id = e.venue_id
  where e.is_public = true;

notify pgrst, 'reload schema';

commit;
