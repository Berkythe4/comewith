-- =============================================================================
-- 144_seed_user_dashboard_prefs.sql  (Strategy rebuild -- carry the hidden set
-- across the singleton -> per-user move)
--
-- 142 added user_dashboard_prefs and Phase 2 switched the Strategy board onto
-- it, but nothing carried the OLD singleton's hidden set over. dashboard_prefs
-- has three cards hidden -- instagram.saves_shares, audience.follower_ticket,
-- youtube.watch_hours, all hand-logged metrics with nothing ever entered -- and
-- without this they would silently reappear for everyone on first load. A card
-- someone deliberately hid coming back is a regression, just a quiet one.
--
-- Seeds every ACTIVE admin. deleted_at is checked because of the 098
-- deactivation contract: a profile with deleted_at set is treated as no-role,
-- so it should not be given prefs here either.
--
-- The old singleton row is left in place, untouched. Nothing reads it now, but
-- it is the only record of what was hidden before this migration ran.
-- =============================================================================
begin;

insert into public.user_dashboard_prefs (user_id, hidden_metric_keys, expanded_categories)
select p.id, coalesce(d.hidden_metric_keys, '{}'::text[]), '{}'::text[]
  from public.profiles p
 cross join (select hidden_metric_keys from public.dashboard_prefs where singleton) d
 where p.role in ('master_admin', 'sub_admin')
   and p.deleted_at is null
    on conflict (user_id) do nothing;

commit;

-- DOWN: delete from public.user_dashboard_prefs;  (the singleton still holds
--       the original hidden set, so nothing is lost by clearing this table.)
