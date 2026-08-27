-- =============================================================================
-- 205_artist_gigs_featured.sql
-- A gig shows on an artist profile if the event faces the public — and this repo
-- says that in TWO flags, not one.
--
-- 204 gated v_artist_gigs on is_public alone. That was half the rule. Per 030's
-- own comment, `is_public` means "surfaces via v_public_events", and both public
-- event views (v_public_events 030, v_public_events_hero 064) also filter
-- `event_date >= current_date`. So is_public is the UPCOMING flag: turning it on
-- for a past event does nothing anywhere else on the site.
--
-- The past-facing flag is `is_featured` — it is what puts an event in Recent
-- Rooms on the homepage, via v_public_recap (061/063/184). An event in Recent
-- Rooms is already named, dated, blurbed and video'd on the front page; its
-- lineup is not a secret.
--
-- So: announced upcoming (is_public) OR publicly recapped (is_featured).
-- Keith's rule, in his words: "everything that is showing in recent rooms".
--
-- Encoded as a rule rather than a one-time flag sync on purpose. Syncing would
-- have meant hand-flipping is_public on today's featured events, and the next
-- event featured in Recent Rooms would have gone missing from its artists'
-- profiles with nothing to explain why.
--
-- Still excluded, and this is the point of the change: an event that is neither
-- announced nor recapped. Today that is the Growth & Networking festivals
-- (Elements, We Belong Here, Hulaween, JunXion) — trips the team took to
-- network, which were being listed as GIGS on artist profiles before 204 — plus
-- private bookings and production jobs.
--
-- Not security_invoker, so it runs as owner and RLS on `events` never applies:
-- this WHERE clause is the whole gate. Columns unchanged, so create-or-replace
-- keeps the anon grant from 065. Consumer: artist.html only.
-- =============================================================================
begin;

create or replace view public.v_artist_gigs as
  select ep.actor_id, e.id as event_id, e.name as event_name, e.event_date,
         v.name as venue_name, ep.role
  from public.event_participants ep
  join public.events e on e.id = ep.event_id and e.deleted_at is null
  left join public.venues v on v.id = e.venue_id
  where e.is_public = true or e.is_featured = true;

notify pgrst, 'reload schema';

commit;
