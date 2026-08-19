-- ============================================================
-- COME WITH — 160 income gets the same treatment as expenses, and the revenue gap
--
-- THE FINDING THAT PROMPTED THIS. Reconciling every event against its money:
-- only Dance Infusion has ANY revenue recorded. Come With 7-11 carries $800 of
-- cost and $0 income; the Crossroads showcase $1,400 and $0; Maxwell House $179
-- and $0. Five booking and production events — Henry Rental, July Jewels,
-- Hulaween, JunXion, b2b Open Decks — have no financials at all.
--
-- So the P&L was not wrong to say Come With earns $550. It was right about the
-- data, and the data is missing every DJ fee, production fee and rental. There is
-- no place in the UI that makes recording one obvious, so nobody has.
--
-- Three things here:
--   1. income gains verified_at / verified_by, so it can have the same one-click
--      review and the same filters the Expenses tab just got.
--   2. A named revenue vocabulary, so "DJ Gig" and "Production Fee" are offered
--      rather than typed differently every time — the same disease vendor_aliases
--      just cured on the payee side.
--   3. v_event_money, which puts cost and revenue for every event side by side
--      and flags the ones spending money while earning none. That list IS the
--      to-do list.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Parity with expenses
-- ---------------------------------------------------------------
alter table public.income add column if not exists verified_at timestamptz;
alter table public.income add column if not exists verified_by uuid references public.profiles(id);

comment on column public.income.verified_at is
  'Set when a human has confirmed this row is correct. Same one-click review the '
  'Expenses tab uses.';

create index if not exists idx_income_verified on public.income(verified_at);
create index if not exists idx_income_event on public.income(event_id);
create index if not exists idx_income_ledger on public.income(ledger);

-- ---------------------------------------------------------------
-- 2. A revenue vocabulary
-- ---------------------------------------------------------------
-- Deliberately a reference table, not a check constraint: Keith will invent a
-- stream we have not thought of, and a constraint would make that a migration.
create table if not exists public.revenue_streams (
  key           text primary key,
  label         text not null,
  applies_to    text,          -- which event series it usually belongs to
  display_order int not null default 100,
  active        boolean not null default true
);

insert into public.revenue_streams (key, label, applies_to, display_order) values
  ('dj_gig',        'DJ Gig fee',        'Bookings',             10),
  ('production',    'Production fee',    'Come With Production', 20),
  ('equipment',     'Equipment rental',  'Bookings',             30),
  ('ticket_sales',  'Ticket sales',      'Come With Parties',    40),
  ('door_split',    'Door split',        'Come With Parties',    50),
  ('sponsorship',   'Sponsorship',       null,                   60),
  ('donation',      'Donation',          'Dance Infusion',       70),
  ('content',       'Content / brand',   'Content Creation',     80),
  ('other',         'Other income',      null,                   99)
on conflict (key) do update set label = excluded.label,
  applies_to = excluded.applies_to, display_order = excluded.display_order;

alter table public.revenue_streams enable row level security;
drop policy if exists "Admins manage revenue streams" on public.revenue_streams;
create policy "Admins manage revenue streams" on public.revenue_streams
  for all using (public.is_admin());
revoke all on public.revenue_streams from anon;

-- ---------------------------------------------------------------
-- 3. Every event's money, side by side
-- ---------------------------------------------------------------
create or replace view public.v_event_money as
select
  e.id as event_id, e.name, e.series, e.event_date, e.status,
  case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end as ledger,
  coalesce(t.amt, 0)                            as ticket_revenue,
  coalesce(s.amt, 0)                            as sponsor_cash,
  coalesce(d.amt, 0)                            as donations,
  coalesce(i.amt, 0)                            as other_income,
  coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) as revenue,
  coalesce(x.amt, 0)                            as expenses,
  coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0)
    - coalesce(x.amt, 0)                        as net,
  -- Spending money and earning none. Past events only: a future booking having no
  -- revenue yet is simply a future booking.
  (coalesce(x.amt,0) > 0
     and coalesce(t.amt,0)+coalesce(s.amt,0)+coalesce(d.amt,0)+coalesce(i.amt,0) = 0
     and e.event_date <= current_date)          as missing_revenue,
  (e.event_date > current_date)                 as upcoming
from public.events e
left join lateral (select sum(amount_paid) amt from public.ticketing where event_id = e.id) t on true
left join lateral (select sum(cash_amount) amt from public.sponsorships where event_id = e.id and status <> 'cancelled') s on true
left join lateral (select sum(amount) amt from public.third_party_donations where event_id = e.id) d on true
left join lateral (select sum(amount) amt from public.income where event_id = e.id and deleted_at is null) i on true
left join lateral (select sum(amount) amt from public.expenses where event_id = e.id and deleted_at is null) x on true
where e.deleted_at is null;

revoke select on public.v_event_money from anon;

commit;

-- DOWN:
--   drop view if exists public.v_event_money;
--   drop table if exists public.revenue_streams;
--   alter table public.income drop column if exists verified_at, drop column if exists verified_by;
