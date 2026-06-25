-- =============================================================================
-- 053_expense_vendor_actor.sql  (additive)
-- expenses.vendor_actor_id — optional link from an expense to a vendor actor
-- (keeps the free-text `vendor` for unmatched Simplifi payees). Admin-only RLS.
-- =============================================================================
begin;
alter table public.expenses add column if not exists vendor_actor_id uuid references public.actors(id) on delete set null;
create index if not exists idx_expenses_vendor_actor on public.expenses(vendor_actor_id) where vendor_actor_id is not null;
commit;
-- DOWN: alter table public.expenses drop column vendor_actor_id;
