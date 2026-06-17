-- =============================================================================
-- 040_guest_kpi_fix.sql  —  Guest KPI fix: fuzzy returning match + mission/business split
--
--  A) v_event_attendance_kpi rebuilt: "returning" now matched by PERSON IDENTITY
--     (normalized full name with a small nickname canon, falling back to email) instead
--     of guest_id/email-exact. So a person who attended two events under different emails
--     or a nickname variant (Liz↔Elizabeth) counts as returning. This is a KPI display
--     calc over a normalized full name — it does NOT merge any guest/actor records.
--  B) v_guest_spend_split: each guest's spend grouped by EVENT TYPE —
--     dance_infusion = MISSION (funds the National MS Society), party = BUSINESS revenue,
--     everything else = OTHER. Type-driven (reads events.type), not hardcoded per event.
--
-- ADDITIVE ONLY: CREATE OR REPLACE views. No tables/money touched. security_invoker.
-- =============================================================================
begin;

create or replace view public.v_event_attendance_kpi with (security_invoker = true) as
with base as (
  select gea.event_id, e.event_date, lower(g.email) as email,
    regexp_replace(
      lower(translate(coalesce(g.full_name,''),
        'áàäâãéèëêíìïîóòöôõúùüûñ','aaaaaeeeeiiiiooooouuuun')),
      '[^a-z0-9 ]', '', 'g') as nname
  from public.guest_event_attendance gea
  join public.guests g on g.id = gea.guest_id and g.deleted_at is null
  join public.events e on e.id = gea.event_id and e.deleted_at is null
),
canon as (
  select event_id, event_date, email,
    case lower(split_part(trim(nname),' ',1))
      when 'liz' then 'elizabeth' when 'beth' then 'elizabeth'
      when 'berky' then 'keith'
      when 'teri' then 'theresa' when 'terri' then 'theresa'
      when 'mike' then 'michael' when 'matt' then 'matthew'
      when 'chris' then 'christopher' when 'dan' then 'daniel'
      when 'zack' then 'zachary' when 'zach' then 'zachary'
      when 'sam' then 'samuel' when 'alex' then 'alexander'
      else split_part(trim(nname),' ',1)
    end || ' ' || nullif(substr(trim(nname), position(' ' in trim(nname)||' ')+1), '') as namekey,
    nname
  from base
),
person as (
  -- identity: canonicalized full name when it has 2+ tokens, else email (avoids single-token collisions)
  select event_id, event_date,
    case when trim(nname) like '% %' then trim(namekey) else coalesce(nullif(trim(nname),''), email) end as pkey
  from canon
),
firstev as (select pkey, min(event_date) as fd from person group by pkey)
select
  e.id as event_id, e.name, e.event_date,
  count(distinct p.pkey)                                          as attendees,
  count(distinct p.pkey) filter (where p.event_date = f.fd)       as new_attendees,
  count(distinct p.pkey) filter (where p.event_date > f.fd)       as returning_attendees,
  round(100.0 * count(distinct p.pkey) filter (where p.event_date > f.fd)
        / nullif(count(distinct p.pkey),0), 1)                    as repeat_pct
from public.events e
join person p on p.event_id = e.id
join firstev f on f.pkey = p.pkey
where e.deleted_at is null
group by e.id, e.name, e.event_date;
revoke all on public.v_event_attendance_kpi from anon;

-- B) mission vs business spend split (type-driven)
create or replace view public.v_guest_spend_split with (security_invoker = true) as
select
  g.id as guest_id,
  coalesce(sum(gea.amount_spent) filter (where e.type = 'dance_infusion'), 0)                                   as mission_spend,
  coalesce(sum(gea.amount_spent) filter (where e.type = 'party'), 0)                                            as business_spend,
  coalesce(sum(gea.amount_spent) filter (where e.type is distinct from 'dance_infusion' and e.type is distinct from 'party'), 0) as other_spend
from public.guests g
left join public.guest_event_attendance gea on gea.guest_id = g.id
left join public.events e on e.id = gea.event_id and e.deleted_at is null
where g.deleted_at is null
group by g.id;
revoke all on public.v_guest_spend_split from anon;

commit;

-- DOWN: restore 038 v_event_attendance_kpi body; drop view v_guest_spend_split.
