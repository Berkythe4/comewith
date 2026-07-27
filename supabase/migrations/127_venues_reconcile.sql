-- 127: reconcile 126 with the venues table that ALREADY EXISTS.
--
-- 126 assumed `venues` was new and ran `create table if not exists` (a no-op) —
-- but venues is the CRM's actor-linked venue table (contact info, in use). So:
--   * DROP the unique(lower(name)) index 126 added — CRM venue names are NOT
--     unique (same name, different city/state), and the Add-venue tool dedupes
--     in code anyway. A DB uniqueness constraint would block legitimate CRM rows.
--   * EXTEND the existing table with the scene/heat-map fields so ONE venue
--     entity serves both the CRM and the future public heat-map.
-- The venues_admin policy 126 added is harmless (guarantees admin write) and stays.

drop index if exists public.venues_name_key;

alter table public.venues
  add column if not exists area       text,
  add column if not exists lat        double precision,   -- heat-map pin (geocode later)
  add column if not exists lng        double precision,
  add column if not exists website    text,
  add column if not exists instagram  text,
  add column if not exists ra_url     text,
  add column if not exists ticket_url text,
  add column if not exists genres     text[],
  add column if not exists is_partner boolean not null default false,
  add column if not exists source     text not null default 'crm',
  add column if not exists ra_id      text;
