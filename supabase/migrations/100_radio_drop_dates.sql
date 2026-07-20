-- =============================================================================
-- 100_radio_drop_dates.sql
-- Scheduled radio drops. Episodes release on a set cadence with dates known in
-- advance — the date lives HERE (the radio tracker), not in `events` (decision
-- 2026-07-20: radio stays out of the events/series system; Keith's "Come With
-- Radio Episode 1" placeholder event is superseded by this and can be deleted).
-- get-station exposes the next upcoming drop (unpublished, dated today+) so the
-- homepage pill + radio hub can tease it; the dashboard sets it per station.
-- =============================================================================
begin;
alter table public.sc_playlists add column if not exists drop_date date;
commit;
-- POST: sc_playlists.drop_date; EP 1's drop set to 2026-07-23 by the apply run.
