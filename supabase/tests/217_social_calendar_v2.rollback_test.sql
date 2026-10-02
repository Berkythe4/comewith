-- =============================================================================
-- 217_social_calendar_v2.sql
-- Social Calendar v2: a simpler post editor, two captions (Claude's draft and
-- the final one Janelle posts), and a narrow Claude connector (social-mcp).
--
-- ADDITIVE. Nothing is dropped or renamed:
--   * social_posts gains account / format / phase / claude_caption /
--     claude_drafted_at / results. The existing `caption` column IS the final
--     caption (relabelled in the UI, not duplicated).
--   * The stage CHECK is WIDENED with 'ready' and 'approved'. Every old value
--     (review / planned / archived) stays legal, so no row and no older client
--     can be refused by it.
--   * social_post_notes gains author_name, so a note written by the connector
--     can say "Claude" - it has no profile row to point author_id at.
--   * connector_log is new: one row per connector call (tool, post, outcome).
--
-- claude_caption is written ONLY by the connector. That is enforced here by a
-- trigger, not by the dashboard: any signed-in caller (auth.uid() not null)
-- who changes it is refused. The connector runs as service role (auth.uid()
-- null) and is the only path that can. "Use this" in the editor copies the
-- draft into `caption`; it never writes claude_caption.
--
-- Backfill (dry-run first, then run): Come With Radio -> phase radio, Dance
-- Infusion -> phase awareness, everything else stays general. account defaults
-- to come_with for every existing row (incl. the Instagram-only ones the sprint
-- names). The two legacy pipeline stages nobody uses are mapped onto the new
-- pipeline (review -> ready, planned -> scheduled); prod holds 0 such rows today.
-- =============================================================================
begin;

-- 1. New post fields ---------------------------------------------------------
alter table public.social_posts add column if not exists account text not null default 'come_with';
alter table public.social_posts add column if not exists format  text;
alter table public.social_posts add column if not exists phase   text not null default 'general';
alter table public.social_posts add column if not exists claude_caption    text;
alter table public.social_posts add column if not exists claude_drafted_at timestamptz;
alter table public.social_posts add column if not exists results jsonb;

alter table public.social_posts drop constraint if exists social_posts_account_check;
alter table public.social_posts add constraint social_posts_account_check
  check (account in ('come_with', 'di', 'collab'));
alter table public.social_posts drop constraint if exists social_posts_format_check;
alter table public.social_posts add constraint social_posts_format_check
  check (format is null or format in ('reel', 'carousel', 'story', 'post'));
alter table public.social_posts drop constraint if exists social_posts_phase_check;
alter table public.social_posts add constraint social_posts_phase_check
  check (phase in ('awareness', 'sponsors', 'radio', 'convert', 'event', 'post', 'general'));
-- {views, likes, shares, saves}, all optional - but it is an object or nothing.
alter table public.social_posts drop constraint if exists social_posts_results_check;
alter table public.social_posts add constraint social_posts_results_check
  check (results is null or jsonb_typeof(results) = 'object');

-- 2. Stage: widen, never narrow ---------------------------------------------
alter table public.social_posts drop constraint if exists social_posts_stage_check;
alter table public.social_posts add constraint social_posts_stage_check
  check (stage in ('idea', 'drafted', 'ready', 'approved', 'scheduled', 'posted',
                   -- legacy, kept legal so nothing on file or in flight is refused
                   'review', 'planned', 'archived'));

-- 3. claude_caption is the connector's alone --------------------------------
create or replace function public.social_posts_claude_guard()
returns trigger
language plpgsql
set search_path = public
as $$
begin
  if auth.uid() is not null then
    if tg_op = 'INSERT' and (new.claude_caption is not null or new.claude_drafted_at is not null) then
      raise exception 'claude_caption is written only by the Claude connector';
    elsif tg_op = 'UPDATE' and (new.claude_caption is distinct from old.claude_caption
                                or new.claude_drafted_at is distinct from old.claude_drafted_at) then
      raise exception 'claude_caption is written only by the Claude connector';
    end if;
  end if;
  return new;
end;
$$;
drop trigger if exists social_posts_claude_guard on public.social_posts;
create trigger social_posts_claude_guard
  before insert or update on public.social_posts
  for each row execute function public.social_posts_claude_guard();

-- 4. Notes can be signed by the connector -----------------------------------
alter table public.social_post_notes add column if not exists author_name text;

-- 5. Connector call log -----------------------------------------------------
create table if not exists public.connector_log (
  id         bigserial primary key,
  at         timestamptz not null default now(),
  connector  text not null default 'social-mcp',
  tool       text not null,
  post_id    uuid,
  ok         boolean not null,
  detail     text
);
create index if not exists connector_log_at on public.connector_log (at desc);
alter table public.connector_log enable row level security;
-- Written by the connector (service role, bypasses RLS). Admins may read it.
-- No write policy on purpose: nobody signed in has a reason to edit the log.
drop policy if exists connector_log_admin_read on public.connector_log;
create policy connector_log_admin_read on public.connector_log for select
  using (public.is_admin());
revoke all on public.connector_log from anon;
revoke all on sequence public.connector_log_id_seq from anon;

-- 6. Backfill ---------------------------------------------------------------
update public.social_posts set phase = 'radio'
 where series = 'Come With Radio' and phase = 'general';
update public.social_posts set phase = 'awareness'
 where series = 'Dance Infusion' and phase = 'general';
update public.social_posts set account = 'come_with'
 where channels = array['instagram']::text[] and account is distinct from 'come_with';
update public.social_posts set stage = 'ready'     where stage = 'review';
update public.social_posts set stage = 'scheduled' where stage = 'planned';

-- (commit removed: test runs inside a rolled-back transaction)

-- =============================================================================
-- POST-APPLY VERIFICATION
--   * select phase, count(*) from social_posts where deleted_at is null group by 1;
--       expect radio 8, awareness 1, general 37 (prod, 2026-10-02)
--   * signed-in update of claude_caption -> refused by the trigger
--   * anon GET social_posts / social_post_notes / connector_log -> []
-- ROLLBACK (data is preserved by the additive shape; this only removes it)
--   drop trigger social_posts_claude_guard on public.social_posts;
--   drop function public.social_posts_claude_guard();
--   drop table public.connector_log;
--   alter table public.social_posts drop column account, drop column format,
--     drop column phase, drop column claude_caption, drop column claude_drafted_at,
--     drop column results;
--   alter table public.social_post_notes drop column author_name;
-- =============================================================================

-- ---- assertions (still inside the transaction) -----------------------------
do $$
declare n_radio int; n_aw int; n_gen int; n_total int; n_bad int;
begin
  select count(*) filter (where phase='radio'), count(*) filter (where phase='awareness'),
         count(*) filter (where phase='general'), count(*)
    into n_radio, n_aw, n_gen, n_total from public.social_posts;
  if n_radio <> (select count(*) from public.social_posts where series='Come With Radio') then raise exception 'radio backfill %', n_radio; end if;
  if n_aw <> (select count(*) from public.social_posts where series='Dance Infusion') then raise exception 'awareness backfill %', n_aw; end if;
  if n_radio + n_aw + n_gen <> n_total then raise exception 'phase coverage'; end if;
  select count(*) into n_bad from public.social_posts where account <> 'come_with';
  if n_bad <> 0 then raise exception 'account backfill %', n_bad; end if;
  -- constraints refuse junk
  begin update public.social_posts set account='nope' where id=(select id from public.social_posts limit 1); raise exception 'ACCOUNT_CHECK_MISSING';
  exception when check_violation then null; end;
  begin update public.social_posts set results='[1]' where id=(select id from public.social_posts limit 1); raise exception 'RESULTS_CHECK_MISSING';
  exception when check_violation then null; end;
  -- new stages are legal, old ones still legal
  update public.social_posts set stage='approved' where id=(select id from public.social_posts limit 1);
  update public.social_posts set stage='archived' where id=(select id from public.social_posts limit 1);
  -- service role (auth.uid() null) CAN write claude_caption
  update public.social_posts set claude_caption='svc ok', claude_drafted_at=now() where id=(select id from public.social_posts limit 1);
end $$;

-- a signed-in admin can NOT
select set_config('request.jwt.claims', '{"sub":"bd71f88d-fd0a-4e3a-a712-7ac958318c8b","role":"authenticated"}', true);
set local role authenticated;
do $$
begin
  begin
    update public.social_posts set claude_caption='sneaky' where id=(select id from public.social_posts where deleted_at is null limit 1);
    raise exception 'GUARD_MISSING';
  exception when others then
    if sqlerrm = 'GUARD_MISSING' then raise; end if;
    if sqlerrm not like '%only by the Claude connector%' then raise exception 'unexpected: %', sqlerrm; end if;
  end;
  -- but an ordinary edit of the final caption still works for them
  update public.social_posts set caption = caption where id=(select id from public.social_posts where deleted_at is null limit 1);
  if not found then raise exception 'admin could not update a post'; end if;
  -- admin can read the connector log (empty), cannot write it
  perform 1 from public.connector_log;
  begin insert into public.connector_log(tool, ok) values ('x', true); raise exception 'LOG_WRITABLE';
  exception when insufficient_privilege then null; when others then if sqlerrm='LOG_WRITABLE' then raise; end if; end;
end $$;
reset role;
select 'MIGRATION_TEST_PASSED' as result;
rollback;
