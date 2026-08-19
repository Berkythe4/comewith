-- =============================================================================
-- 155_gear_watch_run_log.sql
-- A per-run log for Gear Watch, one row per scan.
--
-- Until now the only record of a scan was gear_watch_config.last_status — a
-- single overwritten string ('ok' / 'partial: ebay failed') — plus the raw HTTP
-- response body sitting in net._http_response, which nobody can see from the
-- dashboard. So "did the 8am scan run, and did Reverb actually answer?" could
-- only be answered by an admin running SQL.
--
-- That matters more here than in a normal feature: this scan is the thing
-- standing between a stolen rig being listed and Keith knowing about it. A
-- source that quietly stopped answering a week ago must be VISIBLE, not
-- inferable — the same reason a failed source is never reported as zero results.
--
-- `sources` holds the per-source line verbatim, e.g.
--   {"reverb": "ok — 163 listing(s) fetched",
--    "ebay": "NOT CONFIGURED — set EBAY_CLIENT_ID and EBAY_CLIENT_SECRET",
--    "facebook": "skipped on the schedule to control cost — press Run scan now"}
-- =============================================================================
begin;

create table if not exists public.gear_watch_runs (
  id           uuid primary key default gen_random_uuid(),
  ran_at       timestamptz not null default now(),
  trigger      text not null default 'cron',      -- 'cron' | 'manual'
  sources      jsonb not null default '{}'::jsonb,
  fetched      int  not null default 0,           -- listings pulled, all sources
  matched      int  not null default 0,           -- cleared the gates and scored
  inserted     int  not null default 0,           -- genuinely new
  alerted      int  not null default 0,           -- announced this run
  emailed      boolean not null default false,
  pushed       int  not null default 0,
  duration_ms  int,
  note         text
);

create index if not exists gear_watch_runs_recent_idx
  on public.gear_watch_runs (ran_at desc);

alter table public.gear_watch_runs enable row level security;

drop policy if exists "Admins manage gear watch runs" on public.gear_watch_runs;
create policy "Admins manage gear watch runs"
  on public.gear_watch_runs for all using (public.is_admin());

revoke all on public.gear_watch_runs from anon;

notify pgrst, 'reload schema';
commit;

-- DOWN: drop table if exists public.gear_watch_runs;
