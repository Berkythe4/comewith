-- 134_rename_radio_module.sql
-- Rename the nav label of the radio module from "Artist Radio" to
-- "Come With Radio" (the public brand: radio.html, the homepage lead card and
-- every episode page already say Come With Radio). Key stays 'ra-market' —
-- module keys are referenced by role grants (104), the workflow map and
-- activateTab(), so only the display label changes.
--
-- PRE : module_registry('ra-market').label = 'Artist Radio' (set in 086)
-- POST: module_registry('ra-market').label = 'Come With Radio'

update public.module_registry
   set label = 'Come With Radio'
 where key = 'ra-market';
