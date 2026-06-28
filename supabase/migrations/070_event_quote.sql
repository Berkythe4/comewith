-- =============================================================================
-- 070_event_quote.sql
-- Let the pricing tool attach a saved quote to an event (events.quote jsonb:
-- the quote inputs + computed totals + saved_at). Admin-gated by existing events
-- RLS — no new policy needed.
-- =============================================================================
begin;

alter table public.events
  add column if not exists quote jsonb;

notify pgrst, 'reload schema';
commit;
