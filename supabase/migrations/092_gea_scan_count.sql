-- =============================================================================
-- 092_gea_scan_count.sql
-- Scanned ADMISSIONS per attendance record. attended (091) is per-person, but
-- one customer record can cover a party of barcodes (7-11: Lunaera = 6 scanned
-- admissions on one record) — so person-counts couldn't reconcile with the
-- event's door count. scan_count carries the admissions:
--   null = no scan data covers this record · 0 = scan data existed, never
--   scanned · N>0 = N admissions scanned under this record.
-- attended=true with scan_count null/0 = MANUALLY confirmed attended (the
-- dashboard shows it as "✓ attended*", toggleable on the Customers tab).
-- =============================================================================
begin;

alter table public.guest_event_attendance
  add column if not exists scan_count int;

comment on column public.guest_event_attendance.scan_count is
  'Admissions scanned under this record per RA scan data (party barcodes roll up). null=no scan data, 0=never scanned, N>0=N admissions. attended=true with no scans = manual confirmation (✓* in the dashboard).';

commit;
