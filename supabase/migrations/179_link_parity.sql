-- ============================================================
-- COME WITH — 179 link parity: closing the gaps the audit found
--
-- A full sweep of the 106 tables and 177 foreign keys turned up one theme that
-- explains most of the loose ends:
--
--   SEVERAL TABLES HAVE A "NEEDS A LINK" STATE AND NO WAY TO SAY "THIS
--   GENUINELY HAS NO LINK."
--
-- `expenses.event_na` is the only place that idea exists. Everywhere else an
-- intentional blank is indistinguishable from an oversight, so the queue never
-- empties and people learn to ignore it. Two concrete examples from prod:
--   * 7 expenses have no payee actor because their "vendor" is Food & Beverage,
--     Gas / Transportation, Presents - categories, not payees. There is no payee
--     to link and never will be, but they will be flagged forever.
--   * 23 social posts have no event because they are Content Creation planning
--     slots. Correct, and permanently noisy.
-- 179 adds the missing not-applicable flags. 180 adds a general waiver so a
-- future check can be dismissed with a reason instead of a migration.
--
-- DONATIONS were the least-linked money table in the system: free-text donor
-- names only, while expenses carry vendor_actor_id, income carries actor_id and
-- sponsorships carry actor_id. Nine of the fifteen distinct donor names already
-- match an actor exactly, so repeat donors were invisible. Now linkable, and
-- backfilled where the name is an exact match.
--
-- DELIBERATELY NOT DONE — ticketing, sponsorships and third_party_donations
-- still hard-delete rather than carrying deleted_at like income and expenses.
-- That asymmetry is real but adding the column is the dangerous half of the fix:
-- roughly ten views sum those three tables WITHOUT a deleted_at filter
-- (v_pl_monthly, v_event_money, v_event_summary, the KPI views, 011/018/022/043),
-- so a soft delete would leave the money behind while the row vanished from the
-- UI. Ghost revenue is worse than a hard delete. Recorded here so the next person
-- does not "tidy it up" in one line.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Donations join the actor graph
-- ---------------------------------------------------------------
alter table public.third_party_donations
  add column if not exists actor_id uuid references public.actors(id) on delete set null;
alter table public.third_party_donations
  add column if not exists created_by uuid references public.profiles(id);

comment on column public.third_party_donations.actor_id is
  'The donor as an actor, when they are one. donor_name stays as the name on the '
  'transaction; this is what makes a repeat donor visible as one person.';

create index if not exists idx_tpd_actor on public.third_party_donations(actor_id);

-- Backfill only on an EXACT case-insensitive name match. A fuzzy match here
-- would attach somebody else's money to a real person's record.
update public.third_party_donations d
   set actor_id = a.id
  from public.actors a
 where d.actor_id is null
   and a.deleted_at is null
   and d.donor_name is not null
   and lower(btrim(d.donor_name)) = lower(btrim(a.display_name));

-- ---------------------------------------------------------------
-- 2. The missing "not applicable" states
-- ---------------------------------------------------------------
-- Income gets what expenses has had since 050: a way to say this row is
-- overhead, not an oversight.
alter table public.income add column if not exists event_na boolean not null default false;
comment on column public.income.event_na is
  'True = this income genuinely belongs to no event (interest, retainers, '
  'overhead). Distinguishes a decision from a blank, exactly as expenses.event_na '
  'does. Nothing sums on it - v_pl_monthly still buckets by event_id.';

-- And expenses gets the same for its payee, which is the field with no way out.
alter table public.expenses add column if not exists payee_na boolean not null default false;
comment on column public.expenses.payee_na is
  'True = there is no payee to link. The vendor string on these rows is a '
  'category (Food & Beverage, Gas / Transportation), not a business. Keeps them '
  'out of the "needs a payee" queue without pretending they have one, and they '
  'stay outside 1099 reporting either way.';

-- Social posts that belong to no event and no episode on purpose.
alter table public.social_posts add column if not exists subject_na boolean not null default false;
comment on column public.social_posts.subject_na is
  'True = this post is not about an event or an episode (evergreen / brand / '
  'planning slots). Without it, 23 correct rows sit in the gap list forever.';

create index if not exists idx_income_event_na on public.income(event_na) where event_na;
create index if not exists idx_expenses_payee_na on public.expenses(payee_na) where payee_na;

-- ---------------------------------------------------------------
-- 3. "Who did this" columns become real links
-- ---------------------------------------------------------------
-- Eight columns named *_by / created_by carried no foreign key, so the person
-- they name could be deleted and the reference would rot silently. Added NOT
-- VALID first and validated in the same transaction: if any legacy value points
-- at a profile that no longer exists the whole migration rolls back and says so,
-- rather than half-applying.
do $$
declare
  t record;
begin
  for t in
    select * from (values
      ('event_photos', 'created_by'),
      ('feedback_log', 'created_by'),
      ('metric_snapshots', 'created_by'),
      ('surveys', 'created_by'),
      ('dashboard_prefs', 'updated_by'),
      ('kpi_targets', 'updated_by'),
      ('pricing_config', 'updated_by'),
      ('test_checklist_state', 'updated_by')
    ) as v(tbl, col)
  loop
    if exists (select 1 from information_schema.columns
                where table_schema = 'public' and table_name = t.tbl and column_name = t.col)
       and not exists (select 1 from pg_constraint c
                        where c.conname = t.tbl || '_' || t.col || '_fkey'
                          and c.conrelid = ('public.' || t.tbl)::regclass)
    then
      execute format(
        'alter table public.%I add constraint %I foreign key (%I) references public.profiles(id) on delete set null not valid',
        t.tbl, t.tbl || '_' || t.col || '_fkey', t.col);
      execute format('alter table public.%I validate constraint %I', t.tbl, t.tbl || '_' || t.col || '_fkey');
    end if;
  end loop;
end $$;

-- ---------------------------------------------------------------
-- 4. Event stage, inferred from the status it already has
-- ---------------------------------------------------------------
-- 16 events carry a status and no stage. The two say the same thing at different
-- resolutions, and every pipeline view reads `stage`. Filled from status only
-- where stage is genuinely absent; nothing already set is touched.
update public.events
   set stage = case
         when audited then 'reported'
         when status = 'completed' then 'wrapped'
         when status in ('planning', 'confirmed') then 'planning'
         when status = 'cancelled' then stage
         else 'planning' end
 where deleted_at is null
   and (stage is null or stage = '')
   and status is not null
   and status <> 'cancelled';

commit;

-- DOWN: alter table public.third_party_donations drop column if exists actor_id,
--   drop column if exists created_by; alter table public.income drop column if
--   exists event_na; alter table public.expenses drop column if exists payee_na;
--   alter table public.social_posts drop column if exists subject_na; drop the
--   eight *_fkey constraints. The stage backfill is not reversible - it filled
--   blanks only.
