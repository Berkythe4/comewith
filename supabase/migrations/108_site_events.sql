-- =============================================================================
-- 108_site_events.sql
-- First-party, cookie-less site analytics: pageviews + outbound/CTA clicks on
-- the public pages, so the Strategy board can answer "is anyone visiting, and
-- do they click through?" without a third party.
--
-- PRIVACY / SECURITY SHAPE (deliberate):
--   * No IP address, no user agent, no cookie, no cross-site identifier is
--     stored. Sessions are a random per-TAB id held in sessionStorage, which
--     dies when the tab closes — enough to count visitors, useless for tracking
--     a person. That's what keeps this out of consent-banner territory.
--   * anon gets NOTHING on this table. Writes happen only through the `track`
--     edge function on the service role; reads are admin-only. Same
--     function-only pattern as the radio's get-station (103).
-- =============================================================================
begin;

create table if not exists public.site_events (
  id           bigserial primary key,
  occurred_at  timestamptz not null default now(),
  kind         text not null check (kind in ('view', 'click')),
  -- Page identity: pathname plus only the query keys that name a THING
  -- (?s=<episode>, ?id=<event>). Everything else is dropped client-side and
  -- re-checked in the function, so tracking junk never becomes a page.
  path         text not null,
  -- Referrer HOST only (no path) — enough to tell Instagram from a DM from
  -- Google, without recording the page someone came from.
  referrer     text,
  utm_source   text,
  utm_medium   text,
  utm_campaign text,
  -- click rows only
  link_url     text,
  link_label   text,
  outbound     boolean,
  session_id   text
);

create index if not exists site_events_occurred_idx on public.site_events (occurred_at desc);
create index if not exists site_events_kind_idx     on public.site_events (kind, occurred_at desc);
create index if not exists site_events_path_idx     on public.site_events (path);

alter table public.site_events enable row level security;
drop policy if exists "Admins read site events" on public.site_events;
create policy "Admins read site events" on public.site_events for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.site_events from anon;
revoke all on sequence public.site_events_id_seq from anon;

-- Daily rollup for a traffic screen / sparkline later.
create or replace view public.v_site_traffic_daily as
select occurred_at::date                                            as day,
       count(*) filter (where kind = 'view')                        as views,
       count(distinct session_id) filter (where kind = 'view')      as visitors,
       count(*) filter (where kind = 'click')                       as clicks
  from public.site_events
 group by 1;
revoke select on public.v_site_traffic_daily from anon;

-- ---------------------------------------------------------------------------
-- KPI cards. 30-day rolling so a card reads as "right now" rather than
-- all-time — a lifetime counter only ever goes up and stops meaning anything.
-- ---------------------------------------------------------------------------
create or replace view public.v_kpi_computed as
with di as (
  select k.* from public.v_kpi_dance_infusion k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
pt as (
  select k.* from public.v_kpi_parties k
    join public.events e on e.id = k.event_id where e.status = 'completed'
),
gk as (select * from public.v_guest_kpis limit 1),
se as (select * from public.site_events where occurred_at >= now() - interval '30 days')
select metric_key, value from (values
  ('di.raised_per_event',  (select round(avg(total_raised), 2) from di)),
  ('di.cost_to_raise',     (select round(avg(cost_to_raise_per_dollar), 2) from di)),
  ('di.attendance',        (select round(avg(total_attendance), 0) from di)),
  ('di.to_ms_total',       (select sum(net_pl) from di)),
  ('parties.net_pl',       (select round(avg(net_pl), 2) from pt)),
  ('parties.sell_through', (select round(avg(sell_through_pct), 1) from pt)),
  ('parties.net_pl_total', (select sum(net_pl) from pt)),
  ('audience.subscribers', (select count(*)::numeric from public.subscribers where status = 'subscribed')),
  ('audience.subscribers_come_with', (
     select count(distinct s.id)::numeric from public.subscribers s
       join public.subscriber_segments g on g.subscriber_id = s.id
      where s.status = 'subscribed' and g.segment = 'come_with')),
  ('audience.subscribers_dance_infusion', (
     select count(distinct s.id)::numeric from public.subscribers s
       join public.subscriber_segments g on g.subscriber_id = s.id
      where s.status = 'subscribed' and g.segment = 'dance_infusion')),
  ('radio.playlists_saved', (select count(*)::numeric from public.listener_playlists)),
  ('radio.episode_visits',  (select coalesce(sum(visits), 0)::numeric from public.listener_station_history)),
  ('site.visitors_30d', (select count(distinct session_id)::numeric from se where kind = 'view')),
  ('site.views_30d',    (select count(*)::numeric from se where kind = 'view')),
  ('site.clicks_30d',   (select count(*)::numeric from se where kind = 'click')),
  ('guest.repeat_pct',     (select case when guests_with_attendance > 0 then round(100.0 * repeat_guests / guests_with_attendance, 1) end from gk))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

-- Workstream must stay one of content/audience/parties/dance_infusion — the
-- Strategy board reads WORKSTREAM[ws].color and would throw on a new key.
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('site.visitors_30d', 'audience', 'Site visitors (30d)',   500, 'gte', '', current_date, true),
  ('site.views_30d',    'audience', 'Page views (30d)',     1500, 'gte', '', current_date, true),
  ('site.clicks_30d',   'audience', 'Link clicks (30d)',     150, 'gte', '', current_date, true)
on conflict do nothing;

commit;
