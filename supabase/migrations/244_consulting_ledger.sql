-- =============================================================================
-- 244_consulting_ledger.sql - Keith, 2026-11-16
-- =============================================================================
--   Come With Consulting is a third set of books. It is a SEPARATE SOLE
--   PROPRIETORSHIP (Keith, 2026-11-16), not a line of Come With: all JSD Drip
--   work is Consulting (Lost Lands, Hulaween, the joggers rev-share stake);
--   Frick Frack BlackJack work stays Come With for now.
--
--   1. 'consulting' becomes a legal value of every ledger CHECK (7 tables).
--      Income and expense rows carry their own ledger, so v_pl_monthly already
--      reports a third ledger with no view change. v_cash_position and
--      v_contractor_1099 stay Come With-only on purpose: Consulting has its own
--      bank account and its own 1099 picture.
--   2. invoice_issuers: who an invoice is FROM, per ledger. invoice_settings
--      stays the Come With / default issuer (singleton, unchanged, so nothing
--      reading it breaks). A ledger with a row here overrides it on the
--      document: name, address, prefix, payment details. Master-only, exactly
--      like invoice_settings (bank details at rest; never site_content).
--   3. next_invoice_no(p_ledger): a ledger with its own issuer gets its own
--      gapless yearly sequence (CWC-2026-0001...). Come With / DI keep using
--      invoice_counters and the existing no-arg function, untouched.
--
--   ADDITIVE ONLY. The deployed dashboard does not know 'consulting' and its
--   income/expense forms would save such a row back as 'come_with', so NO
--   consulting rows are written here; data moves only after the UI is live.
-- =============================================================================

begin;

-- 1. ledger checks -------------------------------------------------------------
do $$
declare r record;
begin
  for r in select * from (values
      ('income'), ('expenses'), ('invoices'), ('owner_advances'),
      ('budget_lines'), ('plan_offerings'), ('plan_overrides')) v(t)
  loop
    execute format('alter table public.%I drop constraint %I', r.t, r.t || '_ledger_check');
    execute format('alter table public.%I add constraint %I check (ledger = any (array[''come_with'', ''dance_infusion'', ''consulting'']))', r.t, r.t || '_ledger_check');
  end loop;
end $$;

-- 2. per-ledger issuer ---------------------------------------------------------
create table public.invoice_issuers (
  ledger            text primary key check (ledger in ('dance_infusion', 'consulting')),
  biz_name          text not null,
  biz_legal_name    text,
  biz_address       text,
  biz_email         text,
  biz_phone         text,
  biz_website       text,
  tax_id            text,
  paypal_enabled    boolean not null default false,
  paypal_handle     text,
  paypal_note       text,
  wire_enabled      boolean not null default false,
  wire_bank_name    text,
  wire_beneficiary  text,
  wire_routing      text,
  wire_account      text,
  wire_swift        text,
  wire_bank_address text,
  wire_note         text,
  extra_methods     jsonb not null default '[]'::jsonb check (jsonb_typeof(extra_methods) = 'array'),
  default_notes     text,
  footer_note       text,
  number_prefix     text not null,
  updated_at        timestamptz not null default now(),
  updated_by        uuid references public.profiles(id)
);
comment on table public.invoice_issuers is
  'Who an invoice is FROM, for a ledger other than Come With. Overrides invoice_settings on the document when invoices.ledger matches. Master-only: holds bank details.';

alter table public.invoice_issuers enable row level security;
create policy invoice_issuers_master on public.invoice_issuers
  for all using (public.is_master_admin()) with check (public.is_master_admin());
revoke all on public.invoice_issuers from anon;
grant select, insert, update on public.invoice_issuers to authenticated;
create trigger set_updated_at before update on public.invoice_issuers
  for each row execute function public.handle_updated_at();

-- Bank details are left blank: Keith fills them in under Invoices -> Settings.
insert into public.invoice_issuers (ledger, biz_name, biz_legal_name, number_prefix, default_notes)
values ('consulting', 'Come With Consulting', 'Keith Berkman', 'CWC', 'Thank you for working with Come With Consulting');

-- 3. per-ledger numbering --------------------------------------------------------
create table public.invoice_counters_ledger (
  ledger text not null,
  year   integer not null,
  seq    integer not null default 0,
  primary key (ledger, year)
);
alter table public.invoice_counters_ledger enable row level security;
create policy invoice_counters_ledger_admin on public.invoice_counters_ledger
  for all using (public.is_admin()) with check (public.is_admin());
revoke all on public.invoice_counters_ledger from anon;

create or replace function public.next_invoice_no(p_ledger text)
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
  select number_prefix into pfx from public.invoice_issuers where ledger = p_ledger;
  -- No issuer of its own (Come With, or DI without a row): the shared sequence.
  if pfx is null then
    return public.next_invoice_no();
  end if;
  insert into public.invoice_counters_ledger (ledger, year, seq) values (p_ledger, y, 1)
  on conflict (ledger, year) do update set seq = public.invoice_counters_ledger.seq + 1
  returning seq into n;
  return pfx || '-' || y::text || '-' || lpad(n::text, 4, '0');
end;
$$;
-- EXECUTE is granted to PUBLIC by default; revoking from anon alone is a no-op (§45).
revoke all on function public.next_invoice_no(text) from public, anon;
grant execute on function public.next_invoice_no(text) to authenticated;

commit;
