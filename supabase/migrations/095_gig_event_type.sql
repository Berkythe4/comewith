-- =============================================================================
-- 095_gig_event_type.sql
-- New series 'Bookings' = we (or a collective artist) are booked as talent at
-- someone else's event — distinct from 'Come With Production' (we run their
-- production) so DJ fees don't inflate production-services KPIs. Series is
-- free text (no DDL), but events.type is CHECK-constrained: widen it to allow
-- 'gig', the type the dashboard maps 'Bookings' to (seriesToType).
-- KPI views filter by exact series string, so the new series touches none of
-- the Parties / DI / financial views.
-- =============================================================================
begin;

alter table public.events drop constraint events_type_check;
alter table public.events add constraint events_type_check
  check (type = any (array['party'::text, 'dance_infusion'::text, 'production'::text, 'showcase'::text, 'gig'::text]));

commit;
-- POST-APPLY: insert/update with type='gig' succeeds; existing rows unaffected
-- (constraint is a superset of the old one).
