-- =============================================================================
-- 068_pricing_tool.sql
-- Sales "Pricing" tool: a single-row, admin-only settings store for pricing
-- defaults + per-DJ rate overrides, plus the nav entry.
--
-- The app (assets/pricing-engine.js) carries the industry-standard defaults and
-- deep-merges this row over them — so the seed row can be empty {} and only holds
-- the user's edits/overrides. No anon access (admin tool); RLS = is_admin().
-- =============================================================================
begin;

create table if not exists public.pricing_config (
  id          int primary key default 1 check (id = 1),  -- single-row table
  config      jsonb not null default '{}'::jsonb,
  updated_at  timestamptz not null default now(),
  updated_by  uuid
);

alter table public.pricing_config enable row level security;
drop policy if exists "Admins manage pricing_config" on public.pricing_config;
create policy "Admins manage pricing_config" on public.pricing_config
  for all using (public.is_admin()) with check (public.is_admin());

insert into public.pricing_config (id, config) values (1, '{}'::jsonb)
on conflict (id) do nothing;

-- Nav: Pricing in the Sales group, between Inquiries (10) and Agreements (20).
-- built=true, signed_off=false → master sees it now (review badge); release to
-- staff later from the Team tab. default_roles copied from Agreements for parity.
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
select 'pricing', 'Pricing', 'Sales', 15, true, false, false,
       coalesce((select default_roles from public.module_registry where key = 'agreements'), '{}')
on conflict (key) do update set
  label = excluded.label, nav_group = excluded.nav_group,
  sort_order = excluded.sort_order, built = excluded.built;

notify pgrst, 'reload schema';
commit;
