-- =============================================================================
-- 218_social_posts_restore_planned.sql
-- 217's stage map (planned -> scheduled) also caught 9 SOFT-DELETED posts that
-- were 'planned'. The pre-check counted live rows only (0 planned), so the map
-- was believed to touch nothing. Deleted history should say what it said;
-- 'planned' is still a legal stage. Restored from
-- backups/social_posts_pre217_2026-10-02.json. Live rows are unaffected.
-- =============================================================================
begin;
update public.social_posts set stage = 'planned'
 where stage = 'scheduled' and deleted_at is not null
   and id in ('65e3a6f1-adc8-4dba-a1d8-cd1a81638fa1', '1028bb9a-4cee-48fe-bf1d-b091d1c547c7', '030c5544-77ef-4908-9fb0-f806de156998', '264323de-2a60-48d8-b6df-586de2ba6f43', '946197b7-7efa-4de7-a7da-fa06c0408b24', 'c55909e6-4b81-44bc-af83-2522c916d024', '68b817a3-2c07-4401-b6b7-2669535a887f', '3baf9a5e-0694-4022-b699-b6a70603cefb', '41491f46-4c79-4ade-b65d-a154fd4624dc');
commit;
