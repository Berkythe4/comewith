-- =============================================================================
-- 091_gea_attended.sql
-- "Scanned in" at the person level. ticketing.attended only covers RA ticket
-- barcodes — guest-list scans (comps, performers, door walk-ins) had nowhere to
-- live, so the Customers tab couldn't show who actually walked in. Nullable on
-- purpose: true = scanned in, false = on a list but never scanned,
-- null = unknowable (e.g. Partiful RSVPs — no check-in data in their exports).
-- Set by the event-exports importer from RA scan data; editable by admins.
-- =============================================================================
begin;

alter table public.guest_event_attendance
  add column if not exists attended boolean;

comment on column public.guest_event_attendance.attended is
  'Scanned in at the door (from RA scan data). true=scanned, false=listed but never scanned, null=unknown (no scan data covers this person, e.g. Partiful RSVPs).';

commit;
