-- =============================================================================
-- 245_hulaween_consulting.sql - Keith, 2026-11-16
-- =============================================================================
--   First Come With Consulting work in the books (244 made the ledger).
--   Hulaween, Oct 12-21 2026: Keith worked for JSD Drip. $1,350, PAID IN KIND:
--   applied as Come With Consulting's buy-in to the revenue share on JSD Drip's
--   new joggers line. Keith's call (2026-11-16): book it as income paid in kind
--   - barter is income at fair value - rather than comp it to $0, so the books
--   show what was earned AND what the stake cost.
--     * event (series 'Come With Consulting' -> the dashboard derives the
--       consulting ledger from it), owner JSD Drip, Berky on the lineup
--     * income $1,350 consulting, received, cash_source 'other' (no cash moved)
--     * invoice CWC-2026-0001 (Consulting's own series), paid by one in-kind
--       payment, with created/paid log lines
--   APPLY ONLY AFTER the consulting-ledger dashboard is merged: the deployed
--   one would save a consulting row back as come_with if it were edited.
-- =============================================================================

begin;

do $$
declare
  jsd   uuid := 'a8fe9c38-824e-4aee-8db8-c6124bd933b6';
  berky uuid := '09939421-2095-4151-b943-2d943bf83615';
  keith uuid := 'bd71f88d-fd0a-4e3a-a712-7ac958318c8b';
  ev uuid; inc uuid; inv uuid; n int; no text;
begin
  if not exists (select 1 from public.invoice_issuers where ledger = 'consulting') then raise exception '244 not applied'; end if;
  if exists (select 1 from public.events where slug = 'hulaween-2026' and deleted_at is null) then raise exception 'Hulaween event already exists'; end if;

  insert into public.events (slug, name, event_date, end_date, series, type, status, stage, owner_actor_id, is_public, notes)
  values ('hulaween-2026', 'Hulaween', '2026-10-12', '2026-10-21', 'Come With Consulting', 'gig', 'completed', 'wrapped', jsd, false,
          'Come With Consulting work for JSD Drip, Oct 12-21. $1,350, paid in kind as the buy-in to the JSD Drip joggers revenue share. Expense reimbursement (rental car, tolls, hotel) to follow.')
  returning id into ev;

  insert into public.event_participants (event_id, actor_id, role, roles, notes) values
    (ev, berky, 'crew', array['crew'], 'Come With Consulting, working for JSD Drip'),
    (ev, jsd, 'host', array['host', 'vendor'], 'Client (Come With Consulting)');

  insert into public.income (date, event_id, amount, expected_amount, category, description, status, ledger, actor_id,
                             cash_source, settled_at, created_by)
  values ('2026-10-21', ev, 1350, 1350, 'Other income', 'Hulaween - paid in kind (joggers rev-share buy-in)', 'received', 'consulting', jsd,
          'other', now(), keith)
  returning id into inc;

  insert into public.invoice_counters_ledger (ledger, year, seq) values ('consulting', 2026, 1)
  on conflict (ledger, year) do update set seq = public.invoice_counters_ledger.seq + 1
  returning seq into n;
  if n <> 1 then raise exception 'expected CWC-2026-0001, counter says %', n; end if;
  no := 'CWC-2026-' || lpad(n::text, 4, '0');

  insert into public.invoices (invoice_no, status, bill_to_actor_id, bill_to_name, event_id, issue_date, due_date, terms_days,
                               currency, ledger, notes, pay_paypal, pay_wire, pay_extra, paid_at, created_by)
  values (no, 'paid', jsd, 'JSD Drip', ev, current_date, current_date, 0, 'USD', 'consulting',
          'Paid in kind: applied in full as Come With Consulting''s buy-in to the revenue share on the JSD Drip joggers line. No cash due.',
          false, false, false, now(), keith)
  returning id into inv;

  insert into public.invoice_lines (invoice_id, position, description, detail, qty, unit_price, income_id)
  values (inv, 0, 'Hulaween', 'Work for JSD Drip · Oct 12-21, 2026', 1, 1350, inc);

  insert into public.invoice_payments (invoice_id, paid_on, amount, method, reference, note, created_by)
  values (inv, current_date, 1350, 'other', 'In kind', 'Buy-in to the JSD Drip joggers revenue share (Come With Consulting stake). No cash moved.', keith);

  insert into public.invoice_events (invoice_id, kind, by_profile, detail) values
    (inv, 'created', keith, jsonb_build_object('lines', 1, 'from', 'migration 245')),
    (inv, 'paid', keith, jsonb_build_object('total', 1350, 'paid', 1350, 'method', 'in kind'));
end $$;

commit;
