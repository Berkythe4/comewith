-- ============================================================
-- COME WITH — 151 capital contributions, verification, expense ops
--
-- Three things this adds, all driven by how the P&L is actually used:
--
-- 1. CAPITAL. Money Keith puts into the business was only ever visible as
--    expenses he happened to pay personally. The founding $5,000 was nowhere at
--    all. capital_contributions makes it a real ledger, and v_capital states the
--    whole position: contributed + personally-paid - reimbursed.
--
--    Framing matters: personally-paid spend is INVESTED CAPITAL, not a debt the
--    business owes. Each charge can still be reimbursed individually, which
--    moves it from outstanding to repaid rather than deleting the fact of it.
--
-- 2. VERIFICATION. 247 expenses and no way to tell which have been eyeballed.
--    verified_at is a single timestamp, set by one click — cheap enough that it
--    actually gets used, and filterable so "what still needs a look" is a view
--    rather than a memory exercise.
--
-- 3. TRIAGE. v_expense_health surfaces the four things that make a P&L wrong:
--    no category, no event and not marked overhead, no receipt on a big charge,
--    and same-day/same-amount pairs. That last one is not hypothetical — 66
--    charges existed twice across the two systems before the adopt logic.
--
-- Additive. Every new view is anon-revoked (E1 discipline).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Capital contributions
-- ---------------------------------------------------------------
create table if not exists public.capital_contributions (
  id          uuid primary key default gen_random_uuid(),
  date        date not null,
  amount      numeric(10,2) not null check (amount <> 0),
  kind        text not null default 'cash'
              check (kind in ('cash', 'equipment', 'other')),
  note        text,
  external_ref text,
  created_by  uuid references public.profiles(id),
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now(),
  deleted_at  timestamptz
);

create unique index if not exists uq_capital_external_ref
  on public.capital_contributions(external_ref) where external_ref is not null;
create index if not exists idx_capital_date on public.capital_contributions(date desc);

drop trigger if exists set_updated_at on public.capital_contributions;
create trigger set_updated_at before update on public.capital_contributions
  for each row execute function public.handle_updated_at();

alter table public.capital_contributions enable row level security;
drop policy if exists "Admins can manage capital" on public.capital_contributions;
create policy "Admins can manage capital" on public.capital_contributions
  for all using (public.is_admin());
revoke all on public.capital_contributions from anon;

-- The founding $5,000. Seeded idempotently via external_ref so re-running this
-- migration cannot create a second one.
insert into public.capital_contributions (date, amount, kind, note, external_ref)
values ('2024-12-01', 5000.00, 'cash',
        'Founding capital — the original stake that started Come With', 'seed-founding-5k')
on conflict (external_ref) where external_ref is not null do nothing;

-- ---------------------------------------------------------------
-- 2. Verification + per-charge reimbursement on expenses
-- ---------------------------------------------------------------
alter table public.expenses add column if not exists verified_at timestamptz;
alter table public.expenses add column if not exists verified_by uuid references public.profiles(id);
alter table public.expenses add column if not exists reimbursed_at timestamptz;

comment on column public.expenses.verified_at is
  'Set when a human has confirmed this row is correct and categorised. One click; '
  'null means it still needs a look.';
comment on column public.expenses.reimbursed_at is
  'Only meaningful when funded_by = owner. Set when Keith has paid himself back '
  'for this specific charge; it then stops counting as outstanding capital.';

create index if not exists idx_expenses_verified on public.expenses(verified_at);
create index if not exists idx_expenses_funded_by on public.expenses(funded_by);

-- ---------------------------------------------------------------
-- 3. The capital position, in one row
-- ---------------------------------------------------------------
create or replace view public.v_capital as
with contrib as (
  select coalesce(sum(amount), 0) as amt from public.capital_contributions where deleted_at is null
), personal as (
  select coalesce(sum(amount), 0) as amt,
         coalesce(sum(amount) filter (where reimbursed_at is not null), 0) as repaid
    from public.expenses where deleted_at is null and funded_by = 'owner'
)
select
  contrib.amt                                as contributed,
  personal.amt                               as personally_paid,
  personal.repaid                            as reimbursed,
  contrib.amt + personal.amt                 as invested_gross,
  contrib.amt + personal.amt - personal.repaid as invested_net,
  personal.amt - personal.repaid             as outstanding_reimbursable
from contrib, personal;

revoke select on public.v_capital from anon;

-- ---------------------------------------------------------------
-- 4. What still needs a human
-- ---------------------------------------------------------------
create or replace view public.v_expense_health as
select
  e.id, e.date, e.amount, e.category, e.vendor, e.event_id, e.event_na,
  e.funded_by, e.verified_at, e.receipt_path,
  (e.category is null)                                    as needs_category,
  (e.event_id is null and not e.event_na)                 as needs_event,
  (e.verified_at is null)                                 as needs_review,
  (e.receipt_path is null and e.amount >= 100)            as needs_receipt,
  exists (select 1 from public.expenses d
           where d.id <> e.id and d.deleted_at is null
             and d.date = e.date and d.amount = e.amount)  as possible_duplicate
from public.expenses e
where e.deleted_at is null;

revoke select on public.v_expense_health from anon;

-- ---------------------------------------------------------------
-- 5. Tax-year summary — deductible spend by year and category
-- ---------------------------------------------------------------
-- The `deductible` flag has existed since 005 and has never been surfaced. With
-- a lump sum to offset it is suddenly the most important column on the table.
create or replace view public.v_tax_year as
select
  extract(year from date)::int as tax_year,
  category,
  count(*)                                                    as line_count,
  round(sum(amount), 2)                                       as total,
  round(sum(amount) filter (where deductible), 2)             as deductible_total,
  round(sum(amount) filter (where not deductible), 2)         as non_deductible_total,
  round(sum(amount) filter (where funded_by = 'owner'), 2)    as paid_personally,
  count(*) filter (where receipt_path is null)                as missing_receipts
from public.expenses
where deleted_at is null
group by 1, 2;

revoke select on public.v_tax_year from anon;

commit;

-- DOWN:
--   drop view if exists public.v_tax_year;
--   drop view if exists public.v_expense_health;
--   drop view if exists public.v_capital;
--   alter table public.expenses drop column if exists verified_at,
--     drop column if exists verified_by, drop column if exists reimbursed_at;
--   drop table if exists public.capital_contributions;
