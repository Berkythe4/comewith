-- 133: multi-day events + a new "Growth & Networking" category.
--
-- (a) events.end_date — for events that span more than one day (e.g. a 4-day
--     festival Keith attends to DJ/practice/network). Null = single-day (the norm);
--     event_date stays the start.
-- (b) Extend the type allow-list with 'growth' — industry presence / networking /
--     skill-building. NOT a Come With production and NOT revenue, so it stays out
--     of the party/DI/production KPIs by design (series is free text; no KPI view
--     matches 'Growth & Networking').
alter table public.events add column if not exists end_date date;

alter table public.events drop constraint if exists events_type_check;
alter table public.events add constraint events_type_check
  check (type = any (array['party','dance_infusion','production','showcase','gig','growth']));
