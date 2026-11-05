-- =============================================================================
-- 225_di3_signal_deposit_owner_advance.sql - Signal deposit + Keith's advance
-- =============================================================================
-- Run AFTER 224 (Keith's order, 2026-11-04). DI #3 only.
--
-- Source: Signal invoice dated 09/24/2026, 175 Morgan Ave LLC -> Emergin Leaders
-- Alliance NYC (Keith Berkman): Food & Beverage minimum $6,000, 50% due $3,000.
-- PDF: events/dance-infusion/di-03-2026-11/source-files/Finance/.
--
-- Bluevine (DI account), 2026-11-04:
--   in   $2,000  Keith personal transfer
--   out  $3,000  Signal deposit ($2,000 of Keith's transfer + $1,000 of existing
--                DI Bluevine funds, i.e. Zeffy payouts)
-- Keith's $2,000 is TWO things:
--   $500   his Champion sponsorship (224 created it pending) -> paid
--   $1,500 an OWNER ADVANCE - a new record type. Not revenue, not sponsor cash,
--          not in % to the mission (no money view reads this table). Status
--          'contingent': reimbursable if the $6K minimum is met; if it is not, it
--          LIKELY becomes a donation from Keith - that is Keith's decision, and
--          nothing converts it automatically (resolving requires resolved_at).
--
-- events.doors_time/end_time already read 15:00/22:00 and the "Event start/end
-- time" blocker is already done - asserted, not rewritten.
-- =============================================================================

begin;

create table if not exists public.owner_advances (
  id              uuid primary key default gen_random_uuid(),
  event_id        uuid references public.events(id) on delete set null,
  ledger          text not null default 'come_with' check (ledger in ('come_with', 'dance_infusion')),
  actor_id        uuid references public.actors(id) on delete set null,
  owner_name      text not null,
  amount          numeric(12,2) not null check (amount > 0),
  advanced_on     date not null,
  cash_source     text check (cash_source is null or cash_source in ('paypal', 'bank', 'personal', 'other')),
  account_label   text,
  status          text not null default 'contingent'
                    check (status in ('contingent', 'owed', 'repaid', 'converted_to_donation', 'written_off')),
  condition       text,
  resolution_note text,
  resolved_at     timestamptz,
  resolved_by     uuid,
  notes           text,
  created_by      uuid default auth.uid(),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now(),
  -- open (contingent/owed) <=> unresolved: a conversion is a recorded decision
  check ((status in ('contingent', 'owed')) = (resolved_at is null))
);
create index if not exists owner_advances_event_idx on public.owner_advances (event_id);

drop trigger if exists owner_advances_touch on public.owner_advances;
create trigger owner_advances_touch before update on public.owner_advances
  for each row execute function public.event_control_touch();

alter table public.owner_advances enable row level security;
drop policy if exists "Master manages owner advances" on public.owner_advances;
create policy "Master manages owner advances" on public.owner_advances
  for all using (public.is_master_admin()) with check (public.is_master_admin());
revoke all on public.owner_advances from anon;

do $$
declare
  ev   uuid;
  sp_kb uuid;
  a_kb uuid;
  n int;
begin
  select id into ev from public.events where slug = 'dance-infusion-3' and deleted_at is null;
  if ev is null then raise exception 'DI #3 not found'; end if;
  if exists (select 1 from public.expenses where external_ref like 'signal-inv-2026-09-24:%') then
    raise exception 'Signal deposit already booked - 225 already applied?';
  end if;
  if (select doors_time from public.events where id = ev) <> '15:00' or (select end_time from public.events where id = ev) <> '22:00' then
    raise exception 'event time is not 3-10 PM - look before writing';
  end if;
  if exists (select 1 from public.tasks where event_id = ev and title = 'Event start/end time' and status <> 'done' and deleted_at is null) then
    raise exception 'time blocker still open';
  end if;

  select s.id, s.actor_id into sp_kb, a_kb
  from public.sponsorships s join public.actors a on a.id = s.actor_id
  where s.event_id = ev and a.display_name = 'Keith Berkman' and s.status = 'pending' and s.cash_amount = 500;
  if sp_kb is null then raise exception 'Keith Champion (224) not found pending'; end if;

  -- Signal F&B minimum: deposit paid today from DI Bluevine, balance invoiced
  insert into public.expenses (date, event_id, amount, expected_amount, category, vendor, vendor_actor_id, payment_method,
                               description, ledger, cash_source, funded_by, status, settled_at, due_date, external_ref)
  values
    ('2026-11-04', ev, 3000, 3000, 'Venue', 'Signal NYC', '498dba37-3e7c-42f8-ab42-27c633fc9b6d', 'Bluevine (DI)',
     'Signal F&B minimum - 50% deposit. Invoice 09/24/2026 (175 Morgan Ave LLC -> Emergin Leaders Alliance NYC), $6,000 total. Paid from DI Bluevine: $2,000 of Keith''s 11/4 transfer + $1,000 existing DI funds (Zeffy payouts).',
     'dance_infusion', 'bank', 'business', 'paid', now(), null, 'signal-inv-2026-09-24:deposit'),
    ('2026-11-14', ev, 3000, 3000, 'Venue', 'Signal NYC', '498dba37-3e7c-42f8-ab42-27c633fc9b6d', null,
     'Signal F&B minimum - 50% balance. Invoice 09/24/2026, $6,000 total. Due date TBD.',
     'dance_infusion', null, 'business', 'invoiced', null, null, 'signal-inv-2026-09-24:balance');
  get diagnostics n = row_count;
  if n <> 2 then raise exception 'expected 2 Signal expense rows, got %', n; end if;

  -- $500 of Keith's transfer = his Champion sponsorship
  update public.sponsorships
     set status = 'paid', payment_refs = payment_refs || array['bluevine:2026-11-04:keith-transfer'],
         notes = 'Paid 11/4/2026 via Bluevine (DI): $500 of Keith''s $2,000 personal transfer. Not Zeffy. 0 comps (Keith is on the list as a DJ).'
   where id = sp_kb;

  -- the other $1,500 = owner advance, contingent
  insert into public.owner_advances (event_id, ledger, actor_id, owner_name, amount, advanced_on, cash_source, account_label,
                                     status, condition, notes)
  values (ev, 'dance_infusion', a_kb, 'Keith Berkman', 1500, '2026-11-04', 'bank', 'DI Bluevine', 'contingent',
          'Reimbursable to Keith if the $6,000 Signal F&B minimum is met. If it is not met it likely converts to a donation from Keith - Keith decides; never converted automatically.',
          '$1,500 of Keith''s $2,000 personal transfer to DI Bluevine on 11/4/2026 (the other $500 is his Champion sponsorship). Funded the Signal deposit.');

  insert into public.tasks (event_id, title, description, status, priority, source)
  values (ev, 'Confirm Signal balance due date',
          'Signal F&B minimum $6,000: $3,000 deposit paid 11/4 from DI Bluevine; $3,000 balance invoiced, due date TBD. Set it on the expense once known.',
          'blocked', 'high', 'manual');

  if (select sum(cash_amount) from public.sponsorships where event_id = ev and status = 'paid') <> 1750 then raise exception 'received <> 1750'; end if;
  if (select sum(cash_amount) from public.sponsorships where event_id = ev and status = 'pending') <> 250 then raise exception 'pending <> 250'; end if;
end $$;

commit;

-- DOWN (by hand): delete from owner_advances where event_id = <DI #3>;
--   delete from expenses where external_ref like 'signal-inv-2026-09-24:%';
--   delete from tasks where title = 'Confirm Signal balance due date' and event_id = <DI #3>;
--   update sponsorships set status = 'pending' where <Keith Berkman, DI #3>; drop table owner_advances;
