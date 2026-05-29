-- ============================================================
-- COME WITH — 017 DASHBOARD PREFS  (additive)
-- Single shared row holding which KPI cards are hidden, so Keith and
-- Liz see ONE canonical Strategy layout across devices. Storing the
-- HIDDEN set (default empty) means new metrics appear by default.
-- Admin-only (public.is_admin()).
-- ============================================================
begin;

create table if not exists public.dashboard_prefs (
  singleton          boolean primary key default true check (singleton),
  hidden_metric_keys text[] not null default '{}',
  updated_by         uuid default auth.uid(),
  updated_at         timestamptz not null default now()
);

-- exactly one shared row
insert into public.dashboard_prefs (singleton) values (true)
  on conflict (singleton) do nothing;

alter table public.dashboard_prefs enable row level security;

create policy "Admins can manage dashboard_prefs"
  on public.dashboard_prefs for all using (public.is_admin());

grant usage on schema public to anon, authenticated, service_role;
grant all on all tables in schema public to anon, authenticated, service_role;

commit;
