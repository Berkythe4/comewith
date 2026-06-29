-- =============================================================================
-- 074_equipment_loaded.sql
-- Load-in checkoff: mark each piece of assigned gear as loaded (persists).
-- loaded_at = when it was checked off; null = not loaded yet.
-- =============================================================================
begin;

alter table public.equipment_usage
  add column if not exists loaded_at timestamptz;

notify pgrst, 'reload schema';
commit;
