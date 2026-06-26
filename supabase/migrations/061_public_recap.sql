-- =============================================================================
-- 061_public_recap.sql
-- Admin-curated "Recent Rooms" recap for the public site:
--   * events.is_featured  — admin toggles which events show on the site.
--   * events.youtube_url  — optional recap video (embedded/linked on the card).
--   * events.recap_blurb  — short public caption (separate from internal description).
-- hero photo reuses events.hero_image_path (public 'event-photos' bucket).
-- v_public_recap is anon-readable (public site), exposes only safe fields.
-- =============================================================================
begin;
alter table public.events
  add column if not exists is_featured boolean not null default false,
  add column if not exists youtube_url text,
  add column if not exists recap_blurb text;

create or replace view public.v_public_recap as
select e.id, e.name, e.event_date, v.name as venue_name, e.series, e.type,
       e.hero_image_path, e.youtube_url, e.recap_blurb
from public.events e
left join public.venues v on v.id = e.venue_id
where e.is_featured = true and e.deleted_at is null
order by e.event_date desc;
grant select on public.v_public_recap to anon, authenticated;
commit;
