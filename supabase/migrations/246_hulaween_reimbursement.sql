-- =============================================================================
-- 246_hulaween_reimbursement.sql - Keith, 2026-11-16
-- =============================================================================
--   Hulaween travel Keith paid personally for Come With Consulting, billed back
--   to JSD Drip. Keith: all 12 rental days are JSD Drip's (loading and
--   unloading her gear the whole time).
--     * 5 consulting expenses on the Hulaween event (category Travel, paid,
--       cash_source personal):
--         Enterprise minivan RA 5F00G8, 10/16-10/28   $616.82  (*)
--         TollPass EHI209916712                         $77.83  ($67.93 + 2 x $9.90 fees)
--         TollPass EHI210237131                          $2.33
--         TollPass EHI210613686                          $6.00
--         Hotel, Jacksonville, 10/20                   $100.00
--       (*) the rental agreement's ESTIMATED total; no final receipt yet. If
--           the final differs, edit the expense AND the matching invoice line.
--     * the same 5 as accrued 'Reimbursement' income (consulting), so the P&L
--       nets to zero: the cost is real, and so is JSD Drip owing it back
--     * DRAFT invoice CWC-2026-0002 to JSD Drip, one line per item, $802.98,
--       Net 14, Venmo/Zelle (Consulting's extra methods) only
--   Depends on 245 (the Hulaween event). Apply after the dashboard is merged.
-- =============================================================================

begin;

do $$
declare
  jsd   uuid := 'a8fe9c38-824e-4aee-8db8-c6124bd933b6';
  keith uuid := 'bd71f88d-fd0a-4e3a-a712-7ac958318c8b';
  ev uuid; inv uuid; n int; r record; inc uuid; pos int := 0; total numeric := 0;
begin
  select id into ev from public.events where slug = 'hulaween-2026' and deleted_at is null;
  if ev is null then raise exception 'apply 245 first (no Hulaween event)'; end if;
  if exists (select 1 from public.expenses where event_id = ev and deleted_at is null) then raise exception 'Hulaween already has expenses'; end if;

  insert into public.invoice_counters_ledger (ledger, year, seq) values ('consulting', 2026, 1)
  on conflict (ledger, year) do update set seq = public.invoice_counters_ledger.seq + 1
  returning seq into n;

  insert into public.invoices (invoice_no, status, bill_to_actor_id, bill_to_name, bill_to_email, event_id, issue_date, due_date,
                               terms_days, currency, ledger, notes, pay_paypal, pay_wire, pay_extra, created_by)
  values ('CWC-2026-' || lpad(n::text, 4, '0'), 'draft', jsd, 'JSD Drip', 'jessplendid.inc@gmail.com', ev, current_date, current_date + 14,
          14, 'USD', 'consulting',
          'Reimbursement of Hulaween travel paid by Come With Consulting: the minivan for all 12 days (load-in and load-out of JSD Drip gear), tolls, and one hotel night. Receipts available.',
          false, false, true, keith)
  returning id into inv;

  for r in select * from (values
      (date '2026-10-16', 616.82, 'Enterprise', 'Minivan rental 10/16-10/28 (RA 5F00G8)', 'enterprise:RA-5F00G8',
       'Enterprise minivan, 10/16-10/28 (12 days)', 'RA 5F00G8 - 2024 Chrysler Pacifica'),
      (date '2026-10-28', 77.83, 'Enterprise TollPass', 'Tolls (incl. $19.80 TollPass service fees)', 'enterprise-toll:EHI209916712',
       'Tolls', 'TollPass receipt EHI209916712 ($67.93 tolls + $19.80 service fees)'),
      (date '2026-10-28', 2.33, 'Enterprise TollPass', 'Tolls', 'enterprise-toll:EHI210237131',
       'Tolls', 'TollPass receipt EHI210237131'),
      (date '2026-10-28', 6.00, 'Enterprise TollPass', 'Tolls', 'enterprise-toll:EHI210613686',
       'Tolls', 'TollPass receipt EHI210613686'),
      (date '2026-10-20', 100.00, 'Hotel (Jacksonville)', 'Hotel, Jacksonville, 1 night', 'hotel:jacksonville:2026-10-20',
       'Hotel, Jacksonville', 'One night, Oct 20, 2026')
    ) v(d, amt, vendor, descr, ref, line, detail)
  loop
    insert into public.expenses (date, event_id, amount, expected_amount, category, vendor, description, external_ref,
                                 ledger, cash_source, funded_by, status, settled_at, created_by)
    values (r.d, ev, r.amt, r.amt, 'Travel', r.vendor, r.descr || ' - Hulaween, billed back to JSD Drip', r.ref,
            'consulting', 'personal', 'business', 'paid', now(), keith);
    insert into public.income (date, event_id, amount, expected_amount, category, description, status, ledger, actor_id, external_ref, created_by)
    values (r.d, ev, r.amt, r.amt, 'Reimbursement', 'Reimbursable: ' || r.descr, 'accrued', 'consulting', jsd, r.ref, keith)
    returning id into inc;
    insert into public.invoice_lines (invoice_id, position, description, detail, qty, unit_price, income_id)
    values (inv, pos, r.line, r.detail, 1, r.amt, inc);
    pos := pos + 1; total := total + r.amt;
  end loop;
  if total <> 802.98 then raise exception 'total % <> 802.98', total; end if;

  insert into public.invoice_events (invoice_id, kind, by_profile, detail)
  values (inv, 'created', keith, jsonb_build_object('lines', pos, 'from', 'migration 246'));
end $$;

commit;
