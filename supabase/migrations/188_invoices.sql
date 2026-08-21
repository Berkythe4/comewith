-- ============================================================
-- COME WITH — 188 invoices: billing the revenue that is already on the books
--
-- 161 gave income three states — accrued -> invoiced -> received — and the
-- middle one has never been reachable. Nothing in this database could produce
-- an invoice, so `invoiced` sat in the CHECK constraint as a promise. Every
-- income row on prod today is 'accrued' or 'received'. This migration is what
-- makes the middle state mean something.
--
-- THE ONE RULE THAT KEEPS THE P&L HONEST:
--
--   AN INVOICE IS NOT REVENUE. The income row is the revenue.
--
-- An invoice BILLS income rows that already exist. It never creates them, never
-- sums into the P&L, and no view here is added to any revenue total. If an
-- invoice also booked revenue, every invoiced job would be counted twice — once
-- as the accrual and once as the bill. So:
--
--   income row (accrued)   the money we are owed          <- the P&L counts this
--   invoice + lines        the document asking for it     <- counts nothing
--   invoice_payments       cash arriving against the doc  <- still not revenue
--   income row (received)  settled, when the invoice is   <- the cash date
--                          paid IN FULL
--
-- That last arrow is deliberate. Partial payments live on the INVOICE, because
-- an income row has one `amount` and one `settled_at` and cannot be half
-- settled. A deposit is recorded against the invoice; the income rows flip to
-- received only when the balance reaches zero. See `v_invoice_totals`.
--
-- TOTALS ARE COMPUTED, NEVER STORED. subtotal / discount / tax / total / paid /
-- balance all live in views over the lines and payments. A stored total is a
-- number that can disagree with the rows under it, and on an invoice that
-- disagreement is the difference between what you charged and what you can
-- prove you charged.
--
-- WHERE THE BANK DETAILS GO — and where they must NOT.
--
-- `invoice_settings` holds the PayPal handle and the Bluevine wire details.
-- That table is master_admin only and anon-revoked. It is deliberately NOT
-- `site_content`: site_content is anon-readable by design (it feeds the public
-- site), and an account number in an anon-readable table is a wire-fraud
-- incident, not a config choice. Same reasoning as the Beatport token in
-- CLAUDE.md. Nothing in `public` gets a grant here.
--
-- PUBLIC ACCESS is function-only, like `get-station`. The client-facing invoice
-- page reads through an edge function holding the service role, matched on
-- `public_token`. No anon grant is issued on any table or view in this file.
-- ============================================================
begin;

-- ------------------------------------------------------------------
-- Settings: one row, master-only. Bank details at rest.
-- ------------------------------------------------------------------
create table if not exists public.invoice_settings (
  id                boolean primary key default true check (id),
  biz_name          text not null default 'Come With',
  biz_legal_name    text,
  biz_address       text,
  biz_email         text,
  biz_phone         text,
  biz_website       text default 'comewith.org',
  tax_id            text,
  -- PayPal
  paypal_enabled    boolean not null default false,
  paypal_handle     text,          -- a paypal.me handle, or the account email
  paypal_note       text,
  -- Wire / ACH to the Bluevine business checking account
  wire_enabled      boolean not null default false,
  wire_bank_name    text,
  wire_beneficiary  text,
  wire_routing      text,
  wire_account      text,
  wire_swift        text,
  wire_bank_address text,
  wire_note         text,
  -- Defaults applied to a new invoice
  default_terms_days integer not null default 14,
  default_notes      text,
  footer_note        text,
  number_prefix      text not null default 'CW',
  updated_at        timestamptz not null default now(),
  updated_by        uuid references public.profiles(id)
);

insert into public.invoice_settings (id) values (true) on conflict (id) do nothing;

alter table public.invoice_settings enable row level security;
drop policy if exists invoice_settings_master on public.invoice_settings;
create policy invoice_settings_master on public.invoice_settings
  for all using (public.is_master_admin()) with check (public.is_master_admin());

-- ------------------------------------------------------------------
-- Numbering. Gapless per calendar year, and it must stay that way under
-- concurrency — two people pressing Create at once must not both get 0007.
-- The row lock in the function is what guarantees it; a sequence would not,
-- because a sequence keeps counting after a rollback.
-- ------------------------------------------------------------------
create table if not exists public.invoice_counters (
  year integer primary key,
  seq  integer not null default 0
);
alter table public.invoice_counters enable row level security;
drop policy if exists invoice_counters_admin on public.invoice_counters;
create policy invoice_counters_admin on public.invoice_counters
  for all using (public.is_admin()) with check (public.is_admin());

create or replace function public.next_invoice_no()
returns text
language plpgsql
security definer
set search_path = public
as $$
declare
  y   integer := extract(year from current_date)::integer;
  n   integer;
  pfx text;
begin
  if not public.is_admin() then
    raise exception 'not authorized';
  end if;
  select number_prefix into pfx from public.invoice_settings where id;
  pfx := coalesce(nullif(trim(pfx), ''), 'CW');

  insert into public.invoice_counters (year, seq) values (y, 1)
  on conflict (year) do update set seq = public.invoice_counters.seq + 1
  returning seq into n;

  return pfx || '-' || y::text || '-' || lpad(n::text, 4, '0');
end;
$$;

-- ------------------------------------------------------------------
-- The invoice
-- ------------------------------------------------------------------
create table if not exists public.invoices (
  id               uuid primary key default gen_random_uuid(),
  invoice_no       text not null unique,
  status           text not null default 'draft'
                   check (status in ('draft', 'sent', 'paid', 'void')),
  -- Who is being billed. The actor is the link; the three snapshot columns are
  -- what actually prints. An actor can be renamed, merged or deactivated after
  -- the invoice goes out, and a document that silently rewrites itself years
  -- later is not a record of anything.
  bill_to_actor_id uuid references public.actors(id) on delete set null,
  bill_to_name     text,
  bill_to_email    text,
  bill_to_address  text,
  event_id         uuid references public.events(id) on delete set null,
  issue_date       date not null default current_date,
  due_date         date,
  terms_days       integer,
  currency         text not null default 'USD',
  -- Invoice-level discount, on top of any per-line discount.
  discount_kind    text check (discount_kind in ('amount', 'percent')),
  discount_value   numeric(12,2) not null default 0 check (discount_value >= 0),
  -- Tax is off unless somebody turns it on. A zero-rate tax row on an invoice
  -- that has no tax is a claim, not a blank.
  tax_enabled      boolean not null default false,
  tax_rate         numeric(6,3) not null default 0 check (tax_rate >= 0),
  tax_label        text not null default 'Tax',
  notes            text,
  terms_text       text,
  pay_paypal       boolean not null default true,
  pay_wire         boolean not null default true,
  public_token     uuid not null default gen_random_uuid(),
  sent_at          timestamptz,
  viewed_at        timestamptz,
  paid_at          timestamptz,
  voided_at        timestamptz,
  pdf_path         text,
  ledger           text not null default 'come_with'
                   check (ledger in ('come_with', 'dance_infusion')),
  created_by       uuid references public.profiles(id),
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now(),
  deleted_at       timestamptz
);
create unique index if not exists invoices_public_token_key on public.invoices (public_token);
create index if not exists invoices_status_idx  on public.invoices (status) where deleted_at is null;
create index if not exists invoices_actor_idx   on public.invoices (bill_to_actor_id);
create index if not exists invoices_event_idx   on public.invoices (event_id);

-- ------------------------------------------------------------------
-- Lines. `detail` is the breakdown that prints under the description, so a
-- single billed line can show its parts without pretending each part is a
-- separately priced item.
-- ------------------------------------------------------------------
create table if not exists public.invoice_lines (
  id             uuid primary key default gen_random_uuid(),
  invoice_id     uuid not null references public.invoices(id) on delete cascade,
  position       integer not null default 0,
  description    text not null,
  detail         text,
  qty            numeric(12,3) not null default 1,
  unit_price     numeric(12,2) not null default 0,
  discount_kind  text check (discount_kind in ('amount', 'percent')),
  discount_value numeric(12,2) not null default 0 check (discount_value >= 0),
  taxable        boolean not null default true,
  -- The income row this line bills, when it came from one.
  income_id      uuid references public.income(id) on delete set null,
  created_at     timestamptz not null default now()
);
create index if not exists invoice_lines_invoice_idx on public.invoice_lines (invoice_id, position);

-- An income row belongs to at most ONE live invoice. Without this you can bill
-- the same job twice and both invoices look correct in isolation.
create unique index if not exists invoice_lines_income_once
  on public.invoice_lines (income_id) where income_id is not null;

-- ------------------------------------------------------------------
-- Payments against the invoice. This is where a deposit lives.
-- `income_id` is the row that carried the cash in, when the payment was matched
-- from the bank/PayPal import rather than typed in by hand.
-- ------------------------------------------------------------------
create table if not exists public.invoice_payments (
  id          uuid primary key default gen_random_uuid(),
  invoice_id  uuid not null references public.invoices(id) on delete cascade,
  paid_on     date not null default current_date,
  amount      numeric(12,2) not null check (amount <> 0),
  method      text check (method in ('paypal', 'wire', 'cash', 'check', 'other')),
  reference   text,
  income_id   uuid references public.income(id) on delete set null,
  note        text,
  auto_matched boolean not null default false,
  created_by  uuid references public.profiles(id),
  created_at  timestamptz not null default now()
);
create index if not exists invoice_payments_invoice_idx on public.invoice_payments (invoice_id);

-- ------------------------------------------------------------------
-- RLS. Admin-only on all three, matching every other money table.
-- ------------------------------------------------------------------
alter table public.invoices         enable row level security;
alter table public.invoice_lines    enable row level security;
alter table public.invoice_payments enable row level security;

drop policy if exists invoices_admin         on public.invoices;
drop policy if exists invoice_lines_admin    on public.invoice_lines;
drop policy if exists invoice_payments_admin on public.invoice_payments;

create policy invoices_admin on public.invoices
  for all using (public.is_admin()) with check (public.is_admin());
create policy invoice_lines_admin on public.invoice_lines
  for all using (public.is_admin()) with check (public.is_admin());
create policy invoice_payments_admin on public.invoice_payments
  for all using (public.is_admin()) with check (public.is_admin());

drop trigger if exists set_updated_at on public.invoices;
create trigger set_updated_at before update on public.invoices
  for each row execute function public.handle_updated_at();

-- ------------------------------------------------------------------
-- Computed money. Nothing below is stored.
--
-- Rounding is done ONCE per line and once on the tax, to the cent. Rounding at
-- the end instead lets a 3-line invoice print three amounts that do not add up
-- to the total it prints underneath them.
-- ------------------------------------------------------------------
create or replace view public.v_invoice_line_calc as
select
  l.id,
  l.invoice_id,
  l.position,
  l.description,
  l.detail,
  l.qty,
  l.unit_price,
  l.discount_kind,
  l.discount_value,
  l.taxable,
  l.income_id,
  round(l.qty * l.unit_price, 2) as gross,
  round(
    case
      when l.discount_kind = 'percent' then l.qty * l.unit_price * least(l.discount_value, 100) / 100.0
      when l.discount_kind = 'amount'  then least(l.discount_value, l.qty * l.unit_price)
      else 0
    end, 2) as line_discount,
  round(l.qty * l.unit_price, 2) - round(
    case
      when l.discount_kind = 'percent' then l.qty * l.unit_price * least(l.discount_value, 100) / 100.0
      when l.discount_kind = 'amount'  then least(l.discount_value, l.qty * l.unit_price)
      else 0
    end, 2) as amount
from public.invoice_lines l;

create or replace view public.v_invoice_totals as
with li as (
  select invoice_id,
         count(*)                                             as n_lines,
         coalesce(sum(gross), 0)                              as gross,
         coalesce(sum(line_discount), 0)                      as line_discount,
         coalesce(sum(amount), 0)                             as subtotal,
         coalesce(sum(amount) filter (where taxable), 0)      as taxable_subtotal
  from public.v_invoice_line_calc
  group by invoice_id
),
pay as (
  select invoice_id, coalesce(sum(amount), 0) as paid, max(paid_on) as last_paid_on
  from public.invoice_payments
  group by invoice_id
),
d as (
  select i.id as invoice_id,
         coalesce(li.n_lines, 0)          as n_lines,
         coalesce(li.gross, 0)            as gross,
         coalesce(li.line_discount, 0)    as line_discount,
         coalesce(li.subtotal, 0)         as subtotal,
         coalesce(li.taxable_subtotal, 0) as taxable_subtotal,
         round(
           case
             when i.discount_kind = 'percent'
               then coalesce(li.subtotal, 0) * least(i.discount_value, 100) / 100.0
             when i.discount_kind = 'amount'
               then least(i.discount_value, coalesce(li.subtotal, 0))
             else 0
           end, 2) as invoice_discount,
         coalesce(pay.paid, 0)  as paid,
         pay.last_paid_on,
         i.tax_enabled, i.tax_rate, i.status, i.due_date, i.sent_at
  from public.invoices i
  left join li  on li.invoice_id  = i.id
  left join pay on pay.invoice_id = i.id
)
select
  d.invoice_id,
  d.n_lines,
  d.gross,
  d.line_discount,
  d.subtotal,
  d.invoice_discount,
  -- The invoice-level discount is spread across the taxable portion in the same
  -- proportion it applies to the whole, so turning tax on cannot make a
  -- discounted invoice charge tax on money nobody is paying.
  round(
    greatest(d.taxable_subtotal - case
      when d.subtotal > 0 then d.invoice_discount * (d.taxable_subtotal / d.subtotal)
      else 0 end, 0), 2) as taxable_base,
  case when d.tax_enabled then
    round(greatest(d.taxable_subtotal - case
      when d.subtotal > 0 then d.invoice_discount * (d.taxable_subtotal / d.subtotal)
      else 0 end, 0) * d.tax_rate / 100.0, 2)
  else 0 end as tax,
  d.subtotal - d.invoice_discount + case when d.tax_enabled then
    round(greatest(d.taxable_subtotal - case
      when d.subtotal > 0 then d.invoice_discount * (d.taxable_subtotal / d.subtotal)
      else 0 end, 0) * d.tax_rate / 100.0, 2)
  else 0 end as total,
  d.paid,
  d.last_paid_on,
  d.subtotal - d.invoice_discount + case when d.tax_enabled then
    round(greatest(d.taxable_subtotal - case
      when d.subtotal > 0 then d.invoice_discount * (d.taxable_subtotal / d.subtotal)
      else 0 end, 0) * d.tax_rate / 100.0, 2)
  else 0 end - d.paid as balance,
  -- The state you actually read off the list. `overdue` and `partial` are
  -- DERIVED, never stored - a stored 'overdue' is wrong every midnight.
  case
    when d.status = 'void'  then 'void'
    when d.status = 'draft' then 'draft'
    when d.paid > 0 and d.paid >= (d.subtotal - d.invoice_discount + case when d.tax_enabled then
      round(greatest(d.taxable_subtotal - case
        when d.subtotal > 0 then d.invoice_discount * (d.taxable_subtotal / d.subtotal)
        else 0 end, 0) * d.tax_rate / 100.0, 2) else 0 end) then 'paid'
    when d.paid > 0 then 'partial'
    when d.due_date is not null and d.due_date < current_date then 'overdue'
    else 'sent'
  end as state
from d;

-- The list the dashboard reads. Actor and event names resolved once.
create or replace view public.v_invoices_list as
select
  i.id, i.invoice_no, i.status, i.issue_date, i.due_date, i.currency, i.ledger,
  i.bill_to_actor_id, i.event_id, i.sent_at, i.viewed_at, i.paid_at, i.pdf_path,
  i.created_at, i.public_token,
  coalesce(i.bill_to_name, a.display_name) as bill_to,
  coalesce(i.bill_to_email, a.email)       as bill_to_email,
  e.name        as event_name,
  e.event_date  as event_date,
  t.n_lines, t.subtotal, t.line_discount, t.invoice_discount, t.tax, t.total,
  t.paid, t.balance, t.state, t.last_paid_on
from public.invoices i
left join public.actors  a on a.id = i.bill_to_actor_id
left join public.events  e on e.id = i.event_id
left join public.v_invoice_totals t on t.invoice_id = i.id
where i.deleted_at is null;

-- Which income rows are already billed, for the Income tab's badge. A LEFT JOIN
-- from income would be wrong here: this view answers "billed or not", and rows
-- that are not billed simply are not in it.
create or replace view public.v_income_invoiced as
select l.income_id, i.id as invoice_id, i.invoice_no, i.status, t.state, t.balance
from public.invoice_lines l
join public.invoices i on i.id = l.invoice_id and i.deleted_at is null
left join public.v_invoice_totals t on t.invoice_id = i.id
where l.income_id is not null;

-- ------------------------------------------------------------------
-- Grants. Admin surfaces only. NOTHING to anon - the client-facing page reads
-- through an edge function on the service role, matched on public_token.
-- ------------------------------------------------------------------
revoke select on public.v_invoice_line_calc from anon;
revoke select on public.v_invoice_totals    from anon;
revoke select on public.v_invoices_list     from anon;
revoke select on public.v_income_invoiced   from anon;
revoke select on public.invoices            from anon;
revoke select on public.invoice_lines       from anon;
revoke select on public.invoice_payments    from anon;
revoke select on public.invoice_settings    from anon;
revoke select on public.invoice_counters    from anon;
revoke all    on public.invoice_settings    from anon, authenticated;
grant  all    on public.invoice_settings    to   service_role;
-- authenticated reaches invoice_settings only through the edge function; the
-- account number never leaves the service role.

commit;

-- DOWN:
--   drop view if exists public.v_income_invoiced, public.v_invoices_list,
--                       public.v_invoice_totals, public.v_invoice_line_calc;
--   drop table if exists public.invoice_payments, public.invoice_lines,
--                        public.invoices, public.invoice_counters,
--                        public.invoice_settings;
--   drop function if exists public.next_invoice_no();
