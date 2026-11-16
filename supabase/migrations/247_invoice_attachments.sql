-- =============================================================================
-- 247_invoice_attachments.sql - Keith, 2026-11-16
-- =============================================================================
--   Receipts that go out WITH an invoice email (first use: the Hulaween
--   reimbursement to JSD Drip - Enterprise rental, tolls, hotel).
--     * one row per file attached to an invoice. PRESENCE is the choice: a
--       receipt is attached by adding it and detached by deleting the row
--       (same rule as extra_methods - no "attach: false" rows to flip by accident)
--     * files live in the existing private `receipts` bucket: either an
--       expense's own receipt_path (expense_id set - one upload serves the
--       expense AND the invoice) or a file uploaded straight onto the invoice
--     * invoice-doc's send reads these with the service role and attaches them
--       next to the invoice PDF. Nothing public reads this table.
--   Additive. is_admin(), like invoices / invoice_lines / invoice_payments.
-- =============================================================================

begin;

create table public.invoice_attachments (
  id          uuid primary key default gen_random_uuid(),
  invoice_id  uuid not null references public.invoices(id) on delete cascade,
  path        text not null,
  filename    text not null,
  expense_id  uuid references public.expenses(id) on delete set null,
  created_at  timestamptz not null default now(),
  created_by  uuid references public.profiles(id),
  unique (invoice_id, path)
);
comment on table public.invoice_attachments is
  'Files (in the private receipts bucket) sent with an invoice email. Row present = attached. expense_id set when the file is that expense''s receipt.';
create index invoice_attachments_invoice on public.invoice_attachments(invoice_id);

alter table public.invoice_attachments enable row level security;
create policy invoice_attachments_admin on public.invoice_attachments
  for all using (public.is_admin()) with check (public.is_admin());
revoke all on public.invoice_attachments from anon;
grant select, insert, delete on public.invoice_attachments to authenticated;

commit;
