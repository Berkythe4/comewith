-- =============================================================================
-- 152_gear_watch_krk_single_only.sql
-- Gear Watch: only surface SINGLE KRK monitors.
--
-- Keith's decision 2026-08-19, after every false positive in the first day came
-- from this one target. **Only ONE KRK Rokit 5 was stolen, not the pair**, so a
-- listing offering two is by definition not his — and Rokit 5s are among the
-- most common budget monitors sold, so those listings are constant.
--
-- The gate is the title only, so a description saying "great pair with your
-- other one" doesn't kill a real single listing.
--
-- KNOWN TRADE-OFF, deliberately accepted: 'monitors' plural excludes a genuine
-- SINGLE unit listed as "Studio Monitors". Singular "Studio Monitor" still
-- passes. If real singles start getting missed, drop 'monitors' from the array
-- and keep the rest. Nothing here is a scoring change — these are hard gates, so
-- excluded listings are never stored at all.
--
-- The other three targets are untouched: a PAIR of CDJ-3000s or Wave 8s is the
-- strongest signal we have, because a pair of each is exactly what was taken.
-- =============================================================================
begin;

update public.gear_watch_targets
   set exclude_tokens = array[
         'case', 'cover', 'decal', 'skin', 'stand', 'bag', 'sticker', 'manual',
         'parts', 'broken', 'for parts',
         -- multiples: only one was stolen
         'pair', 'pairs', 'set of', '2x', 'x2', '(2)', 'two', 'both', 'duo',
         'monitors'
       ],
       notes = 'Asset tag S004. ONE monitor stolen, not the pair — multi-unit listings are gated out (152). Rokit 5 is a very common model; expect noise until a serial is known.'
 where label = 'KRK Rokit 5';

notify pgrst, 'reload schema';
commit;

-- DOWN: restore the 146 seed value —
--   update public.gear_watch_targets
--      set exclude_tokens = array['case','cover','stand','bag','sticker','manual','parts','broken','for parts','pair']
--    where label = 'KRK Rokit 5';
