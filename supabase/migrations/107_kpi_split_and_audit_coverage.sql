-- =============================================================================
-- 107_kpi_split_and_audit_coverage.sql
-- 1. Mailing list KPI split by brand (come_with / dance_infusion) + the first
--    radio engagement metrics, all computed live in v_kpi_computed.
-- 2. Audit coverage for the tables the Activity log was blind to — building a
--    radio station, editing the site, sending a campaign, a listener saving a
--    playlist. None of them were audited, so none of that work showed up.
-- FINANCIAL discipline: v_kpi_computed stays anon-revoked (E1 / 015-019 guard).
-- No grants to anon anywhere in here.
-- =============================================================================
begin;

-- ---------------------------------------------------------------------------
-- 1a. Live KPI values. Same shape as 051 — this only ADDS rows to the VALUES
--     list. Brand counts are DISTINCT on subscriber id: a subscriber can hold
--     both brand segments (that's the design), so come_with + dance_infusion
--     deliberately sums to MORE than audience.subscribers. Both respect the
--     global unsubscribe (status = 'subscribed'), same as the overall count.
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
gk as (select * from public.v_guest_kpis limit 1)
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
  -- Radio engagement. Both count SIGNED-IN listeners only — that's all we can
  -- see today; anonymous traffic isn't tracked anywhere yet.
  ('radio.playlists_saved', (select count(*)::numeric from public.listener_playlists)),
  ('radio.episode_visits',  (select coalesce(sum(visits), 0)::numeric from public.listener_station_history)),
  ('guest.repeat_pct',     (select case when guests_with_attendance > 0 then round(100.0 * repeat_guests / guests_with_attendance, 1) end from gk))
) as v(metric_key, value);
revoke select on public.v_kpi_computed from anon;

-- ---------------------------------------------------------------------------
-- 1b. Cards. Workstream MUST be one of content/audience/parties/dance_infusion
--     — the Strategy board reads WORKSTREAM[ws].color and a new key would throw
--     on render. All four of these are audience metrics anyway.
--     Targets are a starting point; edit them in the KPI screen.
-- ---------------------------------------------------------------------------
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('audience.subscribers_come_with',      'audience', 'Mailing list - Come With',       500, 'gte', '', current_date, true),
  ('audience.subscribers_dance_infusion', 'audience', 'Mailing list - Dance Infusion',  500, 'gte', '', current_date, true),
  ('radio.playlists_saved',               'audience', 'Radio playlists saved',           50, 'gte', '', current_date, true),
  ('radio.episode_visits',                'audience', 'Radio episode visits',           500, 'gte', '', current_date, true)
on conflict do nothing;

-- ---------------------------------------------------------------------------
-- 2a. Make the audit function safe for tables whose PK isn't "id".
--     site_content is keyed on `key`; the old body did (new.id)::text, which
--     errors on any table without that column. Behaviour is unchanged for the
--     18 tables already audited — reading id out of the jsonb gives the same
--     value — and DELETE/INSERT never touch the unassigned record.
-- ---------------------------------------------------------------------------
create or replace function public.audit_trigger_function()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  actor_email_val text;
  v_new jsonb := case when tg_op = 'DELETE' then null else to_jsonb(new) end;
  v_old jsonb := case when tg_op = 'INSERT' then null else to_jsonb(old) end;
begin
  select email into actor_email_val from public.profiles where id = auth.uid();

  insert into public.audit_log (table_name, row_id, action, actor_id, actor_email, old_data, new_data)
  values (
    tg_table_name,
    coalesce(v_new->>'id', v_old->>'id', v_new->>'key', v_old->>'key'),
    tg_op,
    auth.uid(),
    actor_email_val,
    v_old,
    v_new
  );
  return coalesce(new, old);
end;
$$;

-- ---------------------------------------------------------------------------
-- 2b. Cover the missing tables. sc_playlists is the headline one — stations are
--     created/renamed/scheduled/published straight from the dashboard with the
--     user's own JWT, so auth.uid() resolves to a real person. Work that runs
--     through a service-role edge function (sc-connect finalize) still logs as
--     "system" — that's inherent to service-role writes, not something a
--     trigger can recover.
--     listener_playlists is deliberately included: it's external-user activity.
-- ---------------------------------------------------------------------------
do $$
declare t text;
begin
  foreach t in array array[
    'sc_playlists', 'sc_playlist_tracks', 'site_content',
    'mailing_campaigns', 'subscribers', 'listener_playlists'
  ] loop
    if to_regclass('public.' || t) is not null then
      execute format('drop trigger if exists audit_%1$s on public.%1$s', t);
      execute format('create trigger audit_%1$s after insert or update or delete on public.%1$s for each row execute function public.audit_trigger_function()', t);
    end if;
  end loop;
end $$;

commit;
