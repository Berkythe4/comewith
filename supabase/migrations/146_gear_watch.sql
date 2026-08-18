-- =============================================================================
-- 146_gear_watch.sql
-- Stolen-gear resale watch. Scans Reverb / eBay / Craigslist three times a day
-- for the DJ rig stolen from a vehicle on 2026-08-1x (grand larceny, NYPD),
-- scores each candidate listing, and surfaces the ones worth a human look.
--
-- Objects:
--   1. gear_watch_targets  — what we hunt (one row per stolen unit/model)
--   2. gear_watch_hits     — candidate listings, deduped, with a score breakdown
--   3. gear_watch_config   — singleton settings (theft date, thresholds, geo)
--   4. gear_watch_kick()   — security-definer caller for pg_cron -> edge function
--   5. cron jobs           — 12:00 / 18:00 / 00:00 UTC = 8am / 2pm / 8pm ET
--   6. module_registry     — 'gearwatch' tab, MASTER ONLY
--
-- Config is a dedicated admin-only table, NOT site_content: site_content is
-- anon-readable and none of this belongs on the public site.
--
-- Safe to re-run. Additive only — no existing object is altered or dropped.
-- =============================================================================
begin;

-- 1. Targets -----------------------------------------------------------------
-- model_tokens   : all lowercase spellings that identify the model in a title
-- exclude_tokens : words that mean the listing is an ACCESSORY for the model,
--                  not the model itself ("case for CDJ-3000", "XDJ-AZ decal").
--                  Without these, every skin, cover and dust lid scores as a hit.
-- serial         : filled in as serials are recovered -> instant top score
-- typical_resale : used-market reference for the price-anomaly metric
create table if not exists public.gear_watch_targets (
  id              uuid primary key default gen_random_uuid(),
  label           text not null,
  make            text,
  model_tokens    text[] not null,
  exclude_tokens  text[] not null default '{}',
  serial          text,
  qty             int not null default 1,
  typical_resale  numeric(10, 2),
  active          boolean not null default true,
  notes           text,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

drop trigger if exists set_updated_at on public.gear_watch_targets;
create trigger set_updated_at
  before update on public.gear_watch_targets
  for each row execute function public.handle_updated_at();

-- 2. Hits --------------------------------------------------------------------
-- (source, listing_id) is the dedupe key: a listing seen on three consecutive
-- scans is ONE row with last_seen_at moving forward, not three rows. alerted_at
-- is what stops the digest re-announcing the same listing every eight hours.
create table if not exists public.gear_watch_hits (
  id              uuid primary key default gen_random_uuid(),
  source          text not null check (source in ('reverb', 'ebay', 'craigslist', 'manual')),
  listing_id      text not null,
  url             text not null,
  title           text not null,
  price           numeric(10, 2),
  currency        text not null default 'USD',
  location        text,
  seller          text,
  seller_feedback int,
  posted_at       timestamptz,
  image_url       text,
  target_id       uuid references public.gear_watch_targets(id) on delete set null,
  score           int not null default 0,
  score_breakdown jsonb not null default '{}'::jsonb,
  status          text not null default 'new'
                    check (status in ('new', 'reviewed', 'dismissed', 'reported', 'recovered')),
  reviewed_by     uuid references auth.users(id) on delete set null,
  reviewed_at     timestamptz,
  review_note     text,
  first_seen_at   timestamptz not null default now(),
  last_seen_at    timestamptz not null default now(),
  alerted_at      timestamptz,
  raw             jsonb,
  unique (source, listing_id)
);

create index if not exists gear_watch_hits_triage_idx
  on public.gear_watch_hits (status, score desc, first_seen_at desc);
create index if not exists gear_watch_hits_seen_idx
  on public.gear_watch_hits (first_seen_at desc);

-- 3. Config (singleton) ------------------------------------------------------
-- theft_date is a HARD GATE in the scorer: a listing that went up before the
-- gear was taken is not the gear. geo_terms are the location strings that count
-- as local. min_score gates the digest; push_score gates the phone buzz.
create table if not exists public.gear_watch_config (
  id            boolean primary key default true check (id),
  theft_date    date,
  -- 65 = "model + one strong signal". Model-match plus recency alone comes to 60
  -- and deliberately stays out of the inbox; model + local + a fresh listing is
  -- 85 and both mails and pushes. Recalibrated 2026-08-18 against real listings.
  min_score     int not null default 65,
  push_score    int not null default 85,
  email_to      text,
  push_user_id  uuid references auth.users(id) on delete set null,
  geo_terms     text[] not null default
    '{new york,ny,nyc,brooklyn,queens,bronx,manhattan,staten island,jersey city,newark,hoboken,yonkers,long island}'::text[],
  enabled       boolean not null default true,
  last_run_at   timestamptz,
  last_status   text,
  updated_at    timestamptz not null default now()
);

drop trigger if exists set_updated_at on public.gear_watch_config;
create trigger set_updated_at
  before update on public.gear_watch_config
  for each row execute function public.handle_updated_at();

insert into public.gear_watch_config (id) values (true) on conflict (id) do nothing;

-- 4. RLS ---------------------------------------------------------------------
alter table public.gear_watch_targets enable row level security;
alter table public.gear_watch_hits    enable row level security;
alter table public.gear_watch_config  enable row level security;

drop policy if exists "Admins manage gear watch targets" on public.gear_watch_targets;
create policy "Admins manage gear watch targets"
  on public.gear_watch_targets for all using (public.is_admin());

drop policy if exists "Admins manage gear watch hits" on public.gear_watch_hits;
create policy "Admins manage gear watch hits"
  on public.gear_watch_hits for all using (public.is_admin());

drop policy if exists "Admins manage gear watch config" on public.gear_watch_config;
create policy "Admins manage gear watch config"
  on public.gear_watch_config for all using (public.is_admin());

-- Belt and braces: this is a police matter, none of it is public.
revoke all on public.gear_watch_targets from anon;
revoke all on public.gear_watch_hits    from anon;
revoke all on public.gear_watch_config  from anon;

-- 5. Seed the targets --------------------------------------------------------
-- Exactly the 10 units Keith confirmed taken on 2026-08-18, collapsed to the 6
-- models worth scanning for. Cases and stands are deliberately NOT seeded as
-- scan targets: they carry no serial, resell for little, and "CDJ case" matches
-- thousands of listings — they would bury the real signal. They stay on the
-- loss schedule; they are just not worth hunting.
insert into public.gear_watch_targets (label, make, model_tokens, exclude_tokens, qty, typical_resale, notes)
values
  ('Pioneer XDJ-AZ', 'Pioneer DJ',
   '{xdj-az,xdjaz,xdj az}',
   '{case,cover,decal,skin,stand,bag,lid,sticker,manual,parts,broken,for parts}',
   1, 2400.00, 'Asset tag D002. Bought direct from Pioneer 2024-12-15.'),
  ('Pioneer CDJ-3000', 'Pioneer DJ',
   '{cdj-3000,cdj3000,cdj 3000}',
   '{case,cover,decal,skin,stand,bag,lid,sticker,manual,parts,broken,for parts}',
   2, 2200.00, 'Asset tags D003 + D004. A PAIR was taken — two in one listing is a strong signal.'),
  ('AlphaTheta Wave 8', 'AlphaTheta',
   '{wave 8,wave-8,wave8}',
   '{case,cover,decal,skin,bag,sticker,manual,parts,broken,for parts}',
   2, 700.00, 'Asset tags S001 + S002. A pair was taken.'),
  ('KRK Rokit 5', 'KRK',
   '{rokit 5,rokit5,rp5}',
   '{case,cover,decal,skin,stand,bag,sticker,manual,parts,broken,for parts,pair}',
   1, 150.00, 'Asset tag S004. Single monitor — a PAIR listing is probably not ours.')
on conflict do nothing;

-- 6. Cron -> edge function ---------------------------------------------------
-- pg_cron cannot mint an admin JWT (the reason scheduled sends were deferred in
-- 014). scan-gear-market therefore accepts a SERVICE-ROLE BEARER, the same door
-- pull-ra-market already opens for exactly this purpose. The key is read from
-- vault at call time and never appears in this file or in git.
--
-- Before the first scheduled run, store the two secrets ONCE (SQL editor or
-- db.py — they are not committed):
--   select vault.create_secret('<service-role key>', 'gear_watch_srk');
--   select vault.create_secret('https://<ref>.supabase.co/functions/v1/scan-gear-market', 'gear_watch_url');
--
-- Until then the job is a documented no-op instead of a failure: it records
-- 'skipped: secrets not set' and returns. A cron job that errors silently every
-- eight hours is worse than one that says why it did nothing.
create or replace function public.gear_watch_kick()
returns void
language plpgsql
security definer
set search_path = public, net, extensions, vault
as $$
declare
  v_url text;
  v_key text;
  v_on  boolean;
begin
  select enabled into v_on from public.gear_watch_config where id;
  if not coalesce(v_on, false) then
    update public.gear_watch_config set last_status = 'skipped: disabled' where id;
    return;
  end if;

  select decrypted_secret into v_url from vault.decrypted_secrets where name = 'gear_watch_url';
  select decrypted_secret into v_key from vault.decrypted_secrets where name = 'gear_watch_srk';

  if v_url is null or v_key is null then
    update public.gear_watch_config
       set last_status = 'skipped: secrets not set (gear_watch_url / gear_watch_srk)'
     where id;
    return;
  end if;

  perform net.http_post(
    url     := v_url,
    headers := jsonb_build_object(
                 'Content-Type',  'application/json',
                 'Authorization', 'Bearer ' || v_key),
    body    := jsonb_build_object('trigger', 'cron'),
    timeout_milliseconds := 120000
  );

  update public.gear_watch_config
     set last_run_at = now(), last_status = 'dispatched'
   where id;
end;
$$;

revoke all on function public.gear_watch_kick() from public, anon, authenticated;

-- Three scans a day, 8am / 2pm / 8pm America/New_York (EDT = UTC-4).
-- Spread across the day because listings get posted and pulled within hours;
-- one nightly scan would miss a same-day flip entirely.
select cron.schedule('gear-watch-morning',   '0 12 * * *', $$select public.gear_watch_kick()$$);
select cron.schedule('gear-watch-afternoon', '0 18 * * *', $$select public.gear_watch_kick()$$);
select cron.schedule('gear-watch-evening',   '0 0  * * *', $$select public.gear_watch_kick()$$);

insert into public.automation_jobs (name, description, cron_expression, edge_function, enabled)
values
  ('gear-watch-morning',   'Stolen-gear resale scan — 8am ET',  '0 12 * * *', 'scan-gear-market', true),
  ('gear-watch-afternoon', 'Stolen-gear resale scan — 2pm ET',  '0 18 * * *', 'scan-gear-market', true),
  ('gear-watch-evening',   'Stolen-gear resale scan — 8pm ET',  '0 0  * * *', 'scan-gear-market', true)
on conflict (name) do update
  set description     = excluded.description,
      cron_expression = excluded.cron_expression,
      edge_function   = excluded.edge_function,
      enabled         = excluded.enabled;

-- 7. Dashboard module --------------------------------------------------------
-- master_only: this is a live police matter, not staff-visible work.
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('gearwatch', 'Gear Watch', 'Operations', 65, true, true, true, '{}')
on conflict (key) do update
  set built = true, signed_off = true, master_only = true;

notify pgrst, 'reload schema';
commit;

-- DOWN:
--   select cron.unschedule('gear-watch-morning');
--   select cron.unschedule('gear-watch-afternoon');
--   select cron.unschedule('gear-watch-evening');
--   drop function if exists public.gear_watch_kick();
--   drop table if exists public.gear_watch_hits;
--   drop table if exists public.gear_watch_targets;
--   drop table if exists public.gear_watch_config;
--   delete from public.module_registry where key = 'gearwatch';
--   delete from public.automation_jobs where name like 'gear-watch-%';
