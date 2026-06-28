-- =============================================================================
-- 069_campaign_cc.sql
-- Add a CC / "also send to" field to email campaigns: a comma-separated list of
-- extra addresses that each receive their own copy when the campaign is sent
-- (sent individually like subscribers — not a literal per-email CC header).
-- Lets you include people who aren't in a subscriber segment (team, sponsor, VIP).
-- =============================================================================
begin;

alter table public.mailing_campaigns
  add column if not exists cc text;

notify pgrst, 'reload schema';
commit;
