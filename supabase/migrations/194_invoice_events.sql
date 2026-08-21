-- ============================================================
-- COME WITH — 194 what happened to this invoice, in order
--
-- Everything an invoice has been through was scattered across five columns and
-- a payments table, and three of those columns can only hold the LATEST value.
-- `sent_at` is one timestamp, so re-sending a reminder overwrote the record of
-- the first send; `viewed_at` likewise. "Did we chase them, and when" was not a
-- question this database could answer.
--
-- invoice_events is an APPEND-ONLY log. Nothing in it is ever updated except
-- delivery_status, which the Resend webhook fills in afterwards on the row it
-- matches by resend_id — the same correlation conversation_messages already
-- uses. So a bounced invoice email shows as bounced on the invoice, rather than
-- looking like it was delivered because the send succeeded.
--
--   created    the draft was raised
--   sent       it went to the client   (carries to/cc/subject + resend_id)
--   viewed     they opened the link    (logged once per open, not just the first)
--   payment    money arrived           (carries amount/method/reference)
--   payment_removed  a payment was deleted or corrected
--   paid       the balance reached zero
--   reopened   it stopped being paid in full
--   voided     withdrawn
--   note       anything typed by hand
--
-- WHY A TABLE AND NOT audit_log. audit_log records column changes for
-- forensics; this is a narrative meant to be read by Keith on the invoice
-- screen, with the amounts and the email addresses already in it. Different
-- audience, different shape.
--
-- BACKFILLED from what already exists, so an invoice raised before today does
-- not open with an empty timeline and look like nothing ever happened to it.
-- ============================================================
begin;

create table if not exists public.invoice_events (
  id          uuid primary key default gen_random_uuid(),
  invoice_id  uuid not null references public.invoices(id) on delete cascade,
  kind        text not null check (kind in (
                'created', 'sent', 'viewed', 'payment', 'payment_removed',
                'paid', 'reopened', 'voided', 'note')),
  at          timestamptz not null default now(),
  by_profile  uuid references public.profiles(id),
  -- Free-form per kind: to/cc/subject for a send, amount/method/reference for a
  -- payment. Kept as jsonb rather than a dozen mostly-null columns.
  detail      jsonb not null default '{}'::jsonb,
  resend_id   text,
  delivery_status text,
  created_at  timestamptz not null default now()
);
create index if not exists invoice_events_invoice_idx on public.invoice_events (invoice_id, at desc);
create index if not exists invoice_events_resend_idx  on public.invoice_events (resend_id) where resend_id is not null;

alter table public.invoice_events enable row level security;
drop policy if exists invoice_events_admin on public.invoice_events;
create policy invoice_events_admin on public.invoice_events
  for all using (public.is_admin()) with check (public.is_admin());

revoke select on public.invoice_events from anon;

-- ---- backfill, newest information last so the order reads correctly ----------
insert into public.invoice_events (invoice_id, kind, at, by_profile, detail)
select i.id, 'created', i.created_at, i.created_by, jsonb_build_object('backfilled', true)
from public.invoices i
where not exists (select 1 from public.invoice_events e where e.invoice_id = i.id and e.kind = 'created');

insert into public.invoice_events (invoice_id, kind, at, detail)
select i.id, 'sent', i.sent_at,
       jsonb_build_object('to', i.bill_to_email, 'backfilled', true)
from public.invoices i
where i.sent_at is not null
  and not exists (select 1 from public.invoice_events e where e.invoice_id = i.id and e.kind = 'sent');

insert into public.invoice_events (invoice_id, kind, at, detail)
select i.id, 'viewed', i.viewed_at, jsonb_build_object('backfilled', true)
from public.invoices i
where i.viewed_at is not null
  and not exists (select 1 from public.invoice_events e where e.invoice_id = i.id and e.kind = 'viewed');

insert into public.invoice_events (invoice_id, kind, at, by_profile, detail)
select p.invoice_id, 'payment', coalesce(p.created_at, p.paid_on::timestamptz), p.created_by,
       jsonb_build_object('amount', p.amount, 'method', p.method, 'reference', p.reference,
                          'paid_on', p.paid_on, 'backfilled', true)
from public.invoice_payments p
where not exists (
  select 1 from public.invoice_events e
  where e.invoice_id = p.invoice_id and e.kind = 'payment'
    and (e.detail->>'payment_id') = p.id::text);

insert into public.invoice_events (invoice_id, kind, at, detail)
select i.id, 'paid', i.paid_at, jsonb_build_object('backfilled', true)
from public.invoices i
where i.paid_at is not null
  and not exists (select 1 from public.invoice_events e where e.invoice_id = i.id and e.kind = 'paid');

insert into public.invoice_events (invoice_id, kind, at, detail)
select i.id, 'voided', i.voided_at, jsonb_build_object('backfilled', true)
from public.invoices i
where i.voided_at is not null
  and not exists (select 1 from public.invoice_events e where e.invoice_id = i.id and e.kind = 'voided');

commit;

-- DOWN: drop table if exists public.invoice_events;
