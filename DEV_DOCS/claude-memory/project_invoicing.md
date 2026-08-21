---
name: project_invoicing
description: Invoicing built 2026-08-21 (migrations 188-194) — an invoice is NOT revenue; own dependency-free PDF writer; Bluevine invoicing is Stripe underneath with no API
metadata:
  type: project
---

Invoicing, built end to end on 2026-08-21. Migrations **188-194**, edge function
**`invoice-doc`** (deployed `--no-verify-jwt`), public page `invoice.html`,
Invoices module in the **Finance** nav group.

**THE RULE: an invoice is NOT revenue.** The income row is the revenue; the
invoice is the document asking for it. Nothing in the feature sums into the P&L.
Sending moves the income rows it bills `accrued -> invoiced`; paying in full
moves them to `received` with the payment's date and method as the cash date and
source. This is what finally made `income.status='invoiced'` (defined in 161)
reachable. See [[project_impact_report_supabase]] for the wider money model.

**Totals are computed, never stored** — and therefore exist TWICE:
`v_invoice_totals` (dashboard) and `computeTotals()` in
`supabase/functions/invoice-doc/template.ts` (PDF + client page). They are
checked against each other by `scripts/check_invoice_sql.sql` and
`scripts/test_invoice.mjs` using the same hand-written numbers. **Change one,
change the other.**

**The PDF writer is ours** — `invoice-doc/pdf.ts`, ~280 lines, no dependencies,
standard-14 Helvetica. Unit-tested from Node (Node 24 strips TS natively; avoid
constructor parameter properties, strip-only mode rejects them). Verify output by
parsing it back with `pypdf`. Layout bugs have all been height estimates that
guessed instead of measuring.

**Bank details live in `invoice_settings`, master-admin only — NEVER
`site_content`**, which is anon-readable. Extra rails (Venmo, Zelle) are rows in
`extra_methods` jsonb, not columns.

**BLUEVINE INVOICING IS STRIPE UNDERNEATH and has no public API** — it cannot be
driven from the dashboard, and adopting it would mean double entry plus losing
the link to the income row. The path to card/ACH is Keith's own Stripe account
(Bluevine settles into it) plus a "Pay by card" method. Don't re-research this.

**Still manual:** matching an imported bank deposit to an open invoice.
`invoice_payments.income_id` / `auto_matched` exist and are unused. Wants a
suggested-match queue confirmed by a human.

Two bugs worth not repeating: revoking a GRANT does not make a table admin-only,
it makes it nobody-only and the error blames the caller (LEARNINGS SS41); and a
screen that changes shared modal chrome must have it reset by `openKpi`, not by
itself (SS42, bit twice).
