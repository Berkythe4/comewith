-- =============================================================================
-- 060_vendor_details.sql
-- Vendors get their own tab + per-vendor metadata. actor_vendor_details holds
-- vendor-specific fields (category, account ref, default payment method, typical
-- use) keyed to the actor. Spend stats are computed live from expenses
-- (vendor_actor_id) — not stored. Categories are auto-seeded from each vendor's
-- dominant expense category. Adds a signed-off "Vendors" module (Partners group).
-- =============================================================================
begin;

create table if not exists public.actor_vendor_details (
  actor_id uuid primary key references public.actors(id) on delete cascade,
  category text,
  account_ref text,                 -- login/account email for software subs, acct # etc.
  default_payment_method text,
  typical_use text,                 -- what we buy / use them for
  updated_at timestamptz not null default now()
);
alter table public.actor_vendor_details enable row level security;
create policy "Vendor details module" on public.actor_vendor_details for all
  using (public.user_can_access_module('vendors') or public.is_master_admin())
  with check (public.user_can_access_module('vendors') or public.is_master_admin());

-- Vendors module (own nav tab).
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('vendors', 'Vendors', 'Partners',
        (select coalesce(max(sort_order),0)+1 from public.module_registry),
        true, true, false, array['operations','full'])
on conflict (key) do update set built = true, signed_off = true, label = 'Vendors';

-- Auto-categorize from the vendor's dominant (highest-spend) expense category.
insert into public.actor_vendor_details (actor_id, category)
select a.id,
  case dom.cat
    when 'Equipment' then 'Equipment & Gear'
    when 'Software' then 'Software & Subscriptions'
    when 'Marketing' then 'Marketing & Ads'
    when 'Travel' then 'Travel'
    when 'Contractors' then 'Talent & Contractors'
    when 'Professional Development' then 'Professional Development'
    when 'Platform fees' then 'Platform & Fees'
    when 'Supplies' then 'Supplies'
    when 'Event / Misc' then 'Venue & Space'
    when 'Entertainment / Event' then 'Venue & Space'
    else 'Other'
  end as category
from public.actors a
join public.actor_roles r on r.actor_id = a.id and r.role = 'vendor'
left join lateral (
  select e.category cat from public.expenses e
  where e.vendor_actor_id = a.id and e.deleted_at is null and e.category is not null
  group by e.category order by sum(e.amount) desc nulls last limit 1
) dom on true
where a.deleted_at is null
on conflict (actor_id) do nothing;

commit;
