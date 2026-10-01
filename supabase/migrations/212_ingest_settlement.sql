-- ============================================================
-- COME WITH — 212 ingest settlement: a queue for what must not be guessed,
--                                    and a record of what each run did
--
-- WHY. `ingest-finance` could only ask one question of an incoming payment:
-- "is there a row with this exact date and amount?" That is blind to the thing
-- Come With actually does — a cost is incurred at the gig and paid weeks later.
-- Two real examples from 2026:
--
--   Henry $150 and Berky $100 were recorded against the 2026-08-16 event, then
--   paid through PayPal on 2026-09-08. The feed could not see the August rows
--   (different date), so it inserted its own. $250 of contractor cost counted
--   twice, in two different months.
--
--   A $357.39 Stripe payout was the NET of a $361.00 production fee already on
--   these books ($361.00 less a $3.61 platform fee). Treated as new money it
--   double-counted the revenue while the cash figure still looked right.
--
-- So the function now SETTLES before it ADDS, and where settling is ambiguous it
-- refuses to choose. That refusal needs somewhere to go, which is this table.
--
-- 177 already built the accrual the settlement writes into: expenses.status
-- (accrued -> invoiced -> paid), settled_at ("when the money actually left"),
-- expected_amount. Nothing here changes those semantics — it lets the importer
-- reach them instead of only ever inserting `paid`.
--
-- Additive: no column dropped, nothing backfilled, both new objects
-- anon-revoked and admin-only (E1).
--
-- NUMBERING: written as 207 against a stale local master and applied to prod
-- under that number, which upserted over the ledger row for 207_link_pages.sql
-- (applied_migrations is keyed on version alone). That record has been restored
-- and this is 212. Both migrations are applied in prod; only the bookkeeping was
-- ever wrong. Pull before you pick a number — MERGE_ROUTINE.md step 0 says so
-- for exactly this reason.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. What the importer would not guess at
-- ---------------------------------------------------------------
create table if not exists public.ingest_queue (
  id             uuid primary key default gen_random_uuid(),
  created_at     timestamptz not null default now(),

  -- the incoming row, verbatim, so the decision can be remade later without
  -- the source file. This is the evidence, not a summary of it.
  external_ref   text,
  kind           text not null default 'expense',
  date           date,
  amount         numeric(10,2),
  vendor         text,
  category       text,
  description    text,
  cash_source    text,
  funded_by      text,
  ledger         text,

  -- WHY it stopped. Never "something was wrong" — the reason drives what the
  -- human is being asked, and the dashboard groups by it.
  --   ambiguous_settlement  more than one row could be this payment
  --   amount_variance       settles a known obligation, for a different amount
  --   vendor_mismatch       amount and timing fit, the payee does not
  --   missing_ref           arrived with no external_ref; cannot be deduped at all
  reason         text not null,
  candidate_ids  uuid[] not null default '{}',
  detail         jsonb,

  status         text not null default 'open',
  resolved_at    timestamptz,
  resolved_by    uuid references public.profiles(id),
  resolution     text,

  constraint ingest_queue_reason_check check (reason in
    ('ambiguous_settlement', 'amount_variance', 'vendor_mismatch', 'missing_ref')),
  constraint ingest_queue_status_check check (status in ('open', 'resolved', 'dismissed')),
  constraint ingest_queue_kind_check check (kind in ('expense', 'income'))
);

comment on table public.ingest_queue is
  'Payments the importer refused to place on its own. An entry here means the '
  'run found more than one plausible answer, or none that fit cleanly — not '
  'that anything failed. Empty is the normal state.';
comment on column public.ingest_queue.candidate_ids is
  'The rows this payment might be settling. Two or more is what ambiguous '
  'means; one with a variance is an amount that moved after it was agreed.';

create index if not exists idx_ingest_queue_status on public.ingest_queue(status)
  where status = 'open';
create index if not exists idx_ingest_queue_ref on public.ingest_queue(external_ref);

-- ---------------------------------------------------------------
-- 2. What each run did — so the site can answer "is this current?"
-- ---------------------------------------------------------------
-- Jennifer runs on Keith's desktop and the site cannot see it. Without this the
-- only way to know whether an import had happened was to notice a number look
-- wrong, which is how the feed sat stale for a month.
create table if not exists public.ingest_runs (
  id           uuid primary key default gen_random_uuid(),
  ran_at       timestamptz not null default now(),
  source       text,
  accepted     integer not null default 0,
  settled      integer not null default 0,
  inserted     integer not null default 0,
  updated      integer not null default 0,
  adopted      integer not null default 0,
  queued       integer not null default 0,
  skipped      integer not null default 0,
  budgets      integer not null default 0,
  report_only  boolean not null default false,
  detail       jsonb,
  problems     text[]
);

comment on table public.ingest_runs is
  'One row per push. `report_only` runs change nothing and exist to show what a '
  'real run WOULD do — the settlement rules reach backwards over months of '
  'history, so they get looked at before they get applied.';

create index if not exists idx_ingest_runs_ran_at on public.ingest_runs(ran_at desc);

-- ---------------------------------------------------------------
-- 3. Open queue, readable by the dashboard
-- ---------------------------------------------------------------
create or replace view public.v_ingest_queue as
select q.id, q.created_at, q.reason, q.kind, q.date, q.amount, q.vendor,
       q.category, q.cash_source, q.external_ref, q.candidate_ids, q.detail,
       cardinality(q.candidate_ids) as candidate_count
  from public.ingest_queue q
 where q.status = 'open'
 order by q.created_at desc;

-- ---------------------------------------------------------------
-- 4. RLS — admin surfaces only, and anon sees nothing financial (E1)
-- ---------------------------------------------------------------
alter table public.ingest_queue enable row level security;
alter table public.ingest_runs  enable row level security;

drop policy if exists ingest_queue_admin on public.ingest_queue;
drop policy if exists ingest_runs_admin  on public.ingest_runs;

create policy ingest_queue_admin on public.ingest_queue
  for all using (public.is_admin()) with check (public.is_admin());
create policy ingest_runs_admin on public.ingest_runs
  for all using (public.is_admin()) with check (public.is_admin());

revoke select on public.ingest_queue from anon;
revoke select on public.ingest_runs  from anon;
revoke select on public.v_ingest_queue from anon;

commit;
