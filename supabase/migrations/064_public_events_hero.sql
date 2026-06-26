-- 064_public_events_hero.sql
-- Expose hero_image_path + series on the public upcoming-events view so an
-- uploaded hero photo shows on the homepage event cards (not just Recent Rooms).
-- Columns appended at the end so create-or-replace is allowed; anon grant is
-- preserved.

create or replace view public.v_public_events as
  select e.name,
         e.event_date,
         v.name as venue_name,
         e.ticket_url,
         e.ticket_label,
         e.series,
         e.hero_image_path
  from public.events e
    left join public.venues v on v.id = e.venue_id
  where e.is_public = true
    and e.event_date >= current_date
    and e.deleted_at is null
    and e.status <> 'cancelled';

notify pgrst, 'reload schema';
