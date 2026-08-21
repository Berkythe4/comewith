-- Does v_invoice_totals agree with computeTotals() in the edge function?
--
--   SBP_REF=yaytdosxfhcqatmhctzk python db.py scripts/check_invoice_sql.sql
--
-- Invoice arithmetic lives in two places on purpose: this view (what the
-- dashboard reads) and template.ts (what the PDF and the client's web page
-- read). They must agree to the cent. These are the rows that pull them apart -
-- a per-line discount AND an invoice-level discount, a non-taxable line, and
-- tax charged on top of a discount.
--
-- Expected, and asserted by scripts/test_invoice.mjs on the JavaScript side:
--   gross=2320 lineDiscount=50 subtotal=2270 invoiceDiscount=227
--   taxableBase=1935 tax=171.73 total=2214.73 paid=800 balance=1414.73
--   state=partial
--
-- The whole thing runs inside BEGIN..ROLLBACK, so it touches prod schema and
-- prod's own numeric behaviour without leaving a row behind - and it never calls
-- next_invoice_no(), so the gapless counter is untouched.

begin;
insert into public.invoices (id, invoice_no, status, issue_date, due_date,
  discount_kind, discount_value, tax_enabled, tax_rate, tax_label)
values ('00000000-0000-0000-0000-0000000000aa', 'XCHECK-1', 'sent', '2026-08-21', '2026-09-04',
  'percent', 10, true, 8.875, 'NY sales tax');
insert into public.invoice_lines (invoice_id, position, description, qty, unit_price, discount_kind, discount_value, taxable) values
 ('00000000-0000-0000-0000-0000000000aa', 0, 'DJ performance - 4 hours', 1, 1200, null, 0, true),
 ('00000000-0000-0000-0000-0000000000aa', 1, 'Sound system rental',      1,  650, 'amount', 50, true),
 ('00000000-0000-0000-0000-0000000000aa', 2, 'Lighting package',         2,  175, null, 0, true),
 ('00000000-0000-0000-0000-0000000000aa', 3, 'Travel and load-in',       1,  120, null, 0, false);
insert into public.invoice_payments (invoice_id, paid_on, amount, method, reference)
values ('00000000-0000-0000-0000-0000000000aa', '2026-08-14', 800, 'wire', 'BLV-88213');

select 'gross='||gross||' lineDiscount='||line_discount||' subtotal='||subtotal
     ||' invoiceDiscount='||invoice_discount||' taxableBase='||taxable_base
     ||' tax='||tax||' total='||total||' paid='||paid||' balance='||balance
     ||' state='||state as sql_totals
from public.v_invoice_totals where invoice_id='00000000-0000-0000-0000-0000000000aa';
rollback;
