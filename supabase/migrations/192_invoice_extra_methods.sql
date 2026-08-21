-- ============================================================
-- COME WITH — 192 pay us however you like
--
-- 188 hardcoded two rails, PayPal and the Bluevine wire, as columns. That was
-- wrong the moment somebody wanted Venmo — and it would have been wrong again
-- for Zelle, Cash App, Wise or whatever comes next, each time costing a
-- migration, a template change and a settings-form change to add a label and a
-- handle.
--
-- `extra_methods` is a jsonb ARRAY of { label, detail, note }:
--
--   [{"label":"Venmo","detail":"@come-with-nyc","note":"Add the invoice number"}]
--
-- Order in the array is display order. Presence is what enables it: a method
-- you do not want offered is deleted, not flagged off, because a disabled row
-- carrying a real account handle is a thing that gets re-enabled by accident.
--
-- The two existing rails keep their own columns. They are not "just another
-- method": wire has six structured fields a client's bank actually needs
-- (beneficiary, routing, account, SWIFT, bank address), and PayPal builds a
-- paypal.me link with the balance in it. Collapsing those into a free-text blob
-- to be tidy would lose both.
--
-- `invoices.pay_extra` mirrors pay_paypal / pay_wire so a single invoice can
-- suppress them, which is the same per-invoice control the other two have.
-- ============================================================
begin;

alter table public.invoice_settings
  add column if not exists extra_methods jsonb not null default '[]'::jsonb;

-- It has to be an ARRAY. A bare object or a string here would render as one
-- broken block on every invoice, and the renderer should not have to guess.
alter table public.invoice_settings drop constraint if exists invoice_settings_extra_methods_is_array;
alter table public.invoice_settings add constraint invoice_settings_extra_methods_is_array
  check (jsonb_typeof(extra_methods) = 'array');

alter table public.invoices
  add column if not exists pay_extra boolean not null default true;

commit;

-- DOWN:
--   alter table public.invoice_settings drop column if exists extra_methods;
--   alter table public.invoices drop column if exists pay_extra;
