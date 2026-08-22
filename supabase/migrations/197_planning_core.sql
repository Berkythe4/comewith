-- ============================================================
-- COME WITH — 197 planning core: offerings, volumes, versions
--
-- Turns this repo from a system that RECORDS money into one that PLANS it. The
-- ask: change how many events we do and watch expected earnings move, forward
-- and backward, for every revenue stream — not just ticketed parties.
--
-- ---------------------------------------------------------------------------
-- WHY AN "OFFERING" AND NOT AN "EVENT TYPE"
-- ---------------------------------------------------------------------------
-- The unit of planning is a repeatable thing you sell: a party, a DJ booking,
-- an equipment rental, a production gig. Each has a price, costs that scale
-- with it, and a count per month. That is the whole model — and it is the SAME
-- model for a SKU (price per unit, cost per unit, how many to order). So the
-- table is `plan_offerings`, not `event_types`, and `creates_event` is a flag
-- rather than an assumption. A rental books no event; a SKU never will.
--
-- Three bases cover everything seen so far:
--   per_unit     flat per occurrence      venue $600 an event, fee $500 a gig
--   per_scale    times the scale driver   $28 a head, $12 a unit
--   pct_revenue  percent of the revenue   platform fee 6%
--
-- `scale` is deliberately abstract and NAMED per offering (`scale_label`):
-- "Paid attendance" for a party, "Units" for a SKU, "Hours" for a rental. That
-- one indirection is what makes this generalise instead of hardcoding events.
--
-- pct_revenue is EXPENSE-ONLY, by constraint. A percent-of-revenue income line
-- would be defined in terms of itself; forbidding it at the schema keeps every
-- view a plain aggregate with no recursion and no evaluation order to get wrong.
--
-- ---------------------------------------------------------------------------
-- WHY THE EXISTING 37 BUDGET ROWS ARE NOT READ BY ANY OF THIS
-- ---------------------------------------------------------------------------
-- budget_lines already holds a hand-built forecast in exactly this shape:
-- "Come With Party #1 (7/11)", "DJ Gig #1", "Equipment Rental #1..#6" — each an
-- income row and an expense row, plus standing Marketing $500 / Software $230.
--
-- But they put the UNIT NAME in `category`, and v_pl_monthly_vs_budget joins
-- plan to actual ON CATEGORY. "DJ Gig #1" matches no P&L category and never
-- will, so every one of those lines has been reporting 100% variance since the
-- day it was written. That is not a rendering bug; the join cannot succeed.
--
-- They are NOT rewritten here — they are Keith's numbers, and deleting them
-- would throw away the only forecast that exists. Instead budget_lines gains
-- `version_id`, and the planner reads ONLY rows that carry one. The 37 legacy
-- rows keep version_id null, stay queryable as history, and cannot double-count
-- against the offerings seeded from them.
--
-- ---------------------------------------------------------------------------
-- WHAT IS DERIVED AND WHAT IS DECLARED (LEARNINGS §26)
-- ---------------------------------------------------------------------------
-- Seeded line AMOUNTS come from rows that already exist: the legacy budget
-- figures, and — for ticket pricing — the real average paid ticket and real
-- average paid attendance computed from `ticketing`. Nothing is invented.
--
-- What cannot be derived is which P&L CATEGORY a legacy lump belongs to: the
-- $1,200 against "Come With Party #1" is one number covering venue, talent and
-- marketing, and splitting it would fabricate evidence that then feeds every
-- variance number downstream. So each seeded line carries `needs_review = true`,
-- and the board shows a model with unreviewed lines as PROVISIONAL rather than
-- quietly forecasting from a guess.
--
-- Additive only. Every new table is admin RLS'd and every new view is
-- anon-revoked (E1 discipline / the 016-017 regression).
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Plan versions — what we said, and when we said it
-- ---------------------------------------------------------------
-- Forecast vs actual is meaningless against a moving target: if the plan is
-- edited in place, "how did we do against forecast" silently means "against the
-- forecast as amended after the fact". So there is always exactly one WORKING
-- version you edit, and publishing FREEZES a copy that is never edited again.
create table if not exists public.plan_versions (
  id             uuid primary key default gen_random_uuid(),
  label          text not null,
  status         text not null default 'working'
                   check (status in ('working', 'published', 'archived')),
  horizon_months smallint not null default 6 check (horizon_months in (6, 9, 12)),
  basis_period   text check (basis_period is null or basis_period ~ '^[0-9]{4}-(0[1-9]|1[0-2])$'),
  published_at   timestamptz,
  notes          text,
  created_by     uuid references public.profiles(id),
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now()
);

-- One working version at a time. A second would make "the plan" ambiguous.
create unique index if not exists uq_plan_versions_one_working
  on public.plan_versions ((status)) where status = 'working';

comment on table public.plan_versions is
  'A round of planning. status=working is the one you edit; publishing freezes a '
  'copy so actual-vs-forecast is measured against what was said at the time.';
comment on column public.plan_versions.horizon_months is
  'How far the rolling forecast runs from basis_period. 6 by default, expandable '
  'to 9 or 12 — rolling, not anchored to a fiscal year.';

-- ---------------------------------------------------------------
-- 2. Offerings — the repeatable unit of business
-- ---------------------------------------------------------------
create table if not exists public.plan_offerings (
  id             uuid primary key default gen_random_uuid(),
  key            text not null unique,
  label          text not null,
  ledger         text not null default 'come_with'
                   check (ledger in ('come_with', 'dance_infusion')),
  creates_event  boolean not null default true,
  event_type     text,     -- events.type stamped when a planned unit is realised
  series         text,     -- events.series — matched EXACTLY by the KPI views
  scale_label    text not null default 'Units',
  default_scale  numeric(12,2) not null default 1 check (default_scale >= 0),
  active         boolean not null default true,
  sort_order     integer not null default 100,
  notes          text,
  created_by     uuid references public.profiles(id),
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now(),
  deleted_at     timestamptz
);

-- An offering that books an event must say what kind, or realising it would
-- have to guess events.type and events.series — and series is matched exactly
-- by every KPI view, so a wrong guess reads as an empty KPI, not as an error.
alter table public.plan_offerings drop constraint if exists plan_offerings_event_shape_check;
alter table public.plan_offerings add constraint plan_offerings_event_shape_check
  check (not creates_event or (event_type is not null and series is not null));

comment on table public.plan_offerings is
  'A repeatable unit of business: a party, a DJ booking, a rental, a production '
  'gig — or, in another business built on this, a SKU. Generic on purpose.';
comment on column public.plan_offerings.scale_label is
  'What one unit is measured in beyond the count itself: "Paid attendance" for a '
  'party, "Units" for a SKU. Named per offering so the UI never says "scale".';

create table if not exists public.plan_offering_lines (
  id           uuid primary key default gen_random_uuid(),
  offering_id  uuid not null references public.plan_offerings(id) on delete cascade,
  direction    text not null check (direction in ('income', 'expense')),
  category     text not null,   -- a REAL P&L category, so variance actually joins
  label        text,
  basis        text not null check (basis in ('per_unit', 'per_scale', 'pct_revenue')),
  amount       numeric(12,4) not null default 0,
  needs_review boolean not null default false,
  sort_order   integer not null default 100,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now(),
  deleted_at   timestamptz
);

-- pct_revenue on an income line would be defined in terms of itself.
alter table public.plan_offering_lines drop constraint if exists plan_offering_lines_pct_is_cost_check;
alter table public.plan_offering_lines add constraint plan_offering_lines_pct_is_cost_check
  check (basis <> 'pct_revenue' or direction = 'expense');

create index if not exists idx_plan_offering_lines_offering
  on public.plan_offering_lines(offering_id) where deleted_at is null;

comment on column public.plan_offering_lines.category is
  'The P&L category this lands on. v_plan_vs_actual joins plan to actual on '
  '(period, category) — putting a unit NAME here is what broke the legacy '
  'budget_lines rows, so this must be a category the P&L actually uses.';
comment on column public.plan_offering_lines.needs_review is
  'Seeded from a lump sum whose category could not be derived. The board shows '
  'the model as provisional until a human confirms it. Never silently forecast '
  'from a guessed split.';

-- ---------------------------------------------------------------
-- 3. Volumes — the lever
-- ---------------------------------------------------------------
create table if not exists public.plan_volumes (
  id           uuid primary key default gen_random_uuid(),
  version_id   uuid not null references public.plan_versions(id) on delete cascade,
  offering_id  uuid not null references public.plan_offerings(id) on delete cascade,
  period       text not null check (period ~ '^[0-9]{4}-(0[1-9]|1[0-2])$'),
  units        numeric(10,2) not null default 0 check (units >= 0),
  scale        numeric(12,2) check (scale is null or scale >= 0),
  notes        text,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now(),
  unique (version_id, offering_id, period)
);

comment on column public.plan_volumes.scale is
  'Per-period override of the offering default (a bigger room in December). '
  'NULL means "use the offering default" — it does not mean zero.';

create index if not exists idx_plan_volumes_version_period
  on public.plan_volumes(version_id, period);

-- ---------------------------------------------------------------
-- 4. Overrides — typing over the model without breaking it
-- ---------------------------------------------------------------
-- The model computes a cell; sometimes you know better ("December venue is
-- comped"). An override REPLACES that cell and is recorded as a decision, so
-- the board can show which numbers are modelled and which were asserted.
create table if not exists public.plan_overrides (
  id          uuid primary key default gen_random_uuid(),
  version_id  uuid not null references public.plan_versions(id) on delete cascade,
  period      text not null check (period ~ '^[0-9]{4}-(0[1-9]|1[0-2])$'),
  ledger      text not null check (ledger in ('come_with', 'dance_infusion')),
  section     text not null check (section in ('revenue', 'direct', 'indirect')),
  category    text not null,
  amount      numeric(12,2) not null,
  reason      text,
  created_by  uuid references public.profiles(id),
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now(),
  unique (version_id, ledger, period, section, category)
);

-- ---------------------------------------------------------------
-- 5. budget_lines joins the planner (standing overhead only)
-- ---------------------------------------------------------------
alter table public.budget_lines add column if not exists version_id uuid
  references public.plan_versions(id) on delete cascade;
alter table public.budget_lines add column if not exists ledger text
  not null default 'come_with';
alter table public.budget_lines drop constraint if exists budget_lines_ledger_check;
alter table public.budget_lines add constraint budget_lines_ledger_check
  check (ledger in ('come_with', 'dance_infusion'));

create index if not exists idx_budget_lines_version on public.budget_lines(version_id);

comment on column public.budget_lines.version_id is
  'Which plan round this line belongs to. NULL = legacy: the 37 hand-built rows '
  'that predate the planner. The planner reads only rows WITH a version, so the '
  'old forecast is preserved as history and cannot double-count against the '
  'offerings seeded from it.';

-- ---------------------------------------------------------------
-- 6. updated_at triggers
-- ---------------------------------------------------------------
drop trigger if exists set_updated_at on public.plan_versions;
create trigger set_updated_at before update on public.plan_versions
  for each row execute function public.handle_updated_at();
drop trigger if exists set_updated_at on public.plan_offerings;
create trigger set_updated_at before update on public.plan_offerings
  for each row execute function public.handle_updated_at();
drop trigger if exists set_updated_at on public.plan_offering_lines;
create trigger set_updated_at before update on public.plan_offering_lines
  for each row execute function public.handle_updated_at();
drop trigger if exists set_updated_at on public.plan_volumes;
create trigger set_updated_at before update on public.plan_volumes
  for each row execute function public.handle_updated_at();
drop trigger if exists set_updated_at on public.plan_overrides;
create trigger set_updated_at before update on public.plan_overrides
  for each row execute function public.handle_updated_at();

-- ---------------------------------------------------------------
-- 7. A published version is frozen
-- ---------------------------------------------------------------
-- Enforced in the database, not in the dashboard. The whole value of a
-- published round is that it cannot be quietly improved after the fact, and a
-- client-side guard is not a guarantee — anyone with a REST token bypasses it.
create or replace function public.plan_frozen_guard()
returns trigger
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
declare
  v_status text;
  v_id     uuid;
begin
  v_id := case tg_op when 'DELETE' then old.version_id else new.version_id end;

  -- budget_lines with no version are legacy rows, outside the planner entirely.
  -- Falling back to the row's own id here would have silently looked the wrong
  -- version up and let a frozen row through, so there is no fallback.
  if v_id is null then
    return case tg_op when 'DELETE' then old else new end;
  end if;

  select status into v_status from public.plan_versions where id = v_id;
  if v_status = 'published' then
    raise exception
      'plan version % is published and cannot be changed — create a new round instead', v_id
      using errcode = 'check_violation';
  end if;
  return case tg_op when 'DELETE' then old else new end;
end;
$$;

comment on function public.plan_frozen_guard() is
  'Refuses writes to any plan row belonging to a published version. Actual vs '
  'forecast only means something if the forecast cannot be edited afterwards.';

drop trigger if exists plan_frozen on public.plan_volumes;
create trigger plan_frozen before insert or update or delete on public.plan_volumes
  for each row execute function public.plan_frozen_guard();
drop trigger if exists plan_frozen on public.plan_overrides;
create trigger plan_frozen before insert or update or delete on public.plan_overrides
  for each row execute function public.plan_frozen_guard();
drop trigger if exists plan_frozen on public.budget_lines;
create trigger plan_frozen before insert or update or delete on public.budget_lines
  for each row execute function public.plan_frozen_guard();

-- Publishing itself is the one transition allowed on a published row, so the
-- guard lives on the child tables only. Demoting a published version back to
-- working would unfreeze history, so that is blocked here.
create or replace function public.plan_version_guard()
returns trigger
language plpgsql
security invoker
set search_path = public, pg_temp
as $$
begin
  if old.status = 'published' and new.status = 'working' then
    raise exception 'a published plan round cannot be reopened — create a new round'
      using errcode = 'check_violation';
  end if;
  if old.status = 'published' and new.status = 'published'
     and (new.label is distinct from old.label
          or new.horizon_months is distinct from old.horizon_months
          or new.basis_period is distinct from old.basis_period) then
    raise exception 'a published plan round is frozen'
      using errcode = 'check_violation';
  end if;
  return new;
end;
$$;

drop trigger if exists plan_version_frozen on public.plan_versions;
create trigger plan_version_frozen before update on public.plan_versions
  for each row execute function public.plan_version_guard();

-- ---------------------------------------------------------------
-- 8. RLS — admin surfaces only
-- ---------------------------------------------------------------
alter table public.plan_versions       enable row level security;
alter table public.plan_offerings      enable row level security;
alter table public.plan_offering_lines enable row level security;
alter table public.plan_volumes        enable row level security;
alter table public.plan_overrides      enable row level security;

drop policy if exists plan_versions_admin       on public.plan_versions;
drop policy if exists plan_offerings_admin      on public.plan_offerings;
drop policy if exists plan_offering_lines_admin on public.plan_offering_lines;
drop policy if exists plan_volumes_admin        on public.plan_volumes;
drop policy if exists plan_overrides_admin      on public.plan_overrides;

create policy plan_versions_admin on public.plan_versions
  for all using (public.is_admin()) with check (public.is_admin());
create policy plan_offerings_admin on public.plan_offerings
  for all using (public.is_admin()) with check (public.is_admin());
create policy plan_offering_lines_admin on public.plan_offering_lines
  for all using (public.is_admin()) with check (public.is_admin());
create policy plan_volumes_admin on public.plan_volumes
  for all using (public.is_admin()) with check (public.is_admin());
create policy plan_overrides_admin on public.plan_overrides
  for all using (public.is_admin()) with check (public.is_admin());

commit;

-- DOWN:
--   drop trigger if exists plan_frozen on public.budget_lines;
--   drop trigger if exists plan_frozen on public.plan_overrides;
--   drop trigger if exists plan_frozen on public.plan_volumes;
--   drop trigger if exists plan_version_frozen on public.plan_versions;
--   drop function if exists public.plan_frozen_guard(), public.plan_version_guard();
--   alter table public.budget_lines drop column if exists version_id;
--   alter table public.budget_lines drop column if exists ledger;
--   drop table if exists public.plan_overrides, public.plan_volumes,
--                        public.plan_offering_lines, public.plan_offerings,
--                        public.plan_versions;
