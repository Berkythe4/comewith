-- =============================================================================
-- 145_event_funnel.sql   (Strategy rebuild -- Phase 3: measure the flywheel)
--
-- The flywheel on the Strategy page asserts Content -> Audience -> Parties ->
-- Dance Infusion and measures none of the arrows. This measures the one arrow
-- that decides whether any of it works: does site traffic turn into tickets?
--
--   site exposure  ->  ticket click  ->  ticket sold  ->  attended
--
-- ATTRIBUTION, and why it is not the obvious thing:
-- The ticket CTA lives on the HOMEPAGE (index.html renders a "Get tickets" link
-- per upcoming event, plus the #nsBtn hero button), NOT on event.html --
-- event.html reads v_public_recap and is a retrospective archive page with no
-- CTA at all. So a ticket click is recorded with path='/', and matching on
-- `path` would attribute nothing. Clicks are instead matched to an event by
-- comparing link_url to events.ticket_url, which is exact and per-event.
--
-- Match ignores the query string: the stored ticket_url carries tracking junk
-- (one is a partiful link with an fbclid), and comparing whole strings would
-- silently miss when that differs by a parameter.
--
-- EXPECT THIS TO READ MOSTLY EMPTY AT FIRST, by construction:
--   * the beacon only started 2026-07-24
--   * only two events have ever had a ticket_url (Knicks G5, Come With 7-11)
--     and both completed on or before 2026-07-11 -- i.e. before the beacon
--   * no UPCOMING event has a ticket_url set yet
-- It fills as soon as an upcoming event gets a ticket link. That is the point:
-- this is a promo tracker for the next event, not a report on old ones. There
-- is no retrospective data to recover, so none is implied.
--
-- All views anon-revoked: 013's default privileges auto-grant new views to
-- anon, and this joins ticketing and attendance.
-- =============================================================================
begin;

-- Top of funnel is SITE-WIDE, not per event: the homepage is where the ticket
-- links live, so its traffic is the shared exposure every upcoming event draws
-- from. Splitting it per event would be inventing a number.
create or replace view public.v_site_exposure_30d as
select count(*) filter (where kind = 'view' and path = '/')                          as home_views,
       count(distinct session_id) filter (where kind = 'view' and path = '/')        as home_visitors,
       count(*) filter (where kind = 'view')                                         as all_views,
       count(distinct session_id) filter (where kind = 'view')                       as all_visitors
  from public.site_events
 where occurred_at >= now() - interval '30 days';
revoke select on public.v_site_exposure_30d from anon;

create or replace view public.v_event_funnel as
with ev as (
  select e.id, e.name, e.event_date, e.series, e.status, e.ticket_url,
         e.total_attendance, v.capacity
    from public.events e
    left join public.venues v on v.id = e.venue_id
   where e.deleted_at is null
), clicks as (
  -- Query string stripped on BOTH sides before comparing.
  select split_part(s.link_url, '?', 1) as base_url,
         count(*)                       as clicks,
         count(distinct s.session_id)   as clickers,
         min(s.occurred_at)             as first_click,
         max(s.occurred_at)             as last_click
    from public.site_events s
   where s.kind = 'click' and s.link_url is not null
   group by 1
), views as (
  -- The archive page, where one exists. Separate from the funnel proper: a
  -- recap-page view is interest AFTER the fact, not a step toward a ticket.
  select substring(path from '\?id=([0-9a-fA-F-]{36})')::uuid as event_id,
         count(*)                     as archive_views,
         count(distinct session_id)   as archive_visitors
    from public.site_events
   where kind = 'view' and path like '/event.html?id=%'
   group by 1
), sold as (
  select event_id, tickets_sold from public.v_event_summary
)
select ev.id as event_id, ev.name, ev.event_date, ev.series, ev.status,
       ev.ticket_url is not null and ev.ticket_url <> '' as has_ticket_link,
       coalesce(c.clicks, 0)            as ticket_clicks,
       coalesce(c.clickers, 0)          as ticket_clickers,
       c.first_click, c.last_click,
       coalesce(v.archive_views, 0)     as archive_views,
       coalesce(v.archive_visitors, 0)  as archive_visitors,
       s.tickets_sold,
       ev.capacity,
       ev.total_attendance,
       -- Two conversion rates, both null rather than 0 when the denominator is
       -- missing: a rate of "0%" reads as failure, "no data" reads as no data.
       case when coalesce(c.clickers, 0) > 0 and s.tickets_sold is not null
            then round(100.0 * s.tickets_sold / c.clickers, 1) end as sold_per_clicker_pct,
       case when ev.capacity > 0 and s.tickets_sold is not null
            then round(100.0 * s.tickets_sold / ev.capacity, 1) end as sell_through_pct
  from ev
  left join clicks c on ev.ticket_url is not null and ev.ticket_url <> ''
                    and c.base_url = split_part(ev.ticket_url, '?', 1)
  left join views  v on v.event_id = ev.id
  left join sold   s on s.event_id = ev.id;
revoke select on public.v_event_funnel from anon;

commit;

-- DOWN: drop view public.v_event_funnel; drop view public.v_site_exposure_30d;
