-- =============================================================================
-- 138_notes_assignment.sql
-- Ownership on the Notes page (public.feedback_log), so a note can be claimed by
-- exactly one teammate and two people stop building the same thing twice.
--
-- WHY profiles and not actors: tasks assign to public.actors because a task can
-- land on a DJ, a vendor or a venue contact. Notes is the internal build log —
-- its readers are the five login users — and the notification bell (121) keys on
-- the auth user id, so assigning to a profile means "notify" is a straight
-- lookup instead of an actor -> user mapping that does not exist.
--
-- ONE assignee, not a link table: the whole point is a single owner. A note with
-- three assignees is the double-work problem wearing a hat.
--
-- Additive throughout. No policy change — "Notes module access" is FOR ALL and
-- already covers new columns. No grants — 013's ALTER DEFAULT PRIVILEGES has it.
-- =============================================================================
begin;

alter table public.feedback_log
  add column if not exists assigned_to uuid references public.profiles(id) on delete set null;

alter table public.feedback_log
  add column if not exists assigned_at timestamptz;

-- "What is on my plate" is the question this table now has to answer fast.
create index if not exists feedback_log_assigned_idx
  on public.feedback_log (assigned_to, status)
  where assigned_to is not null;

-- Stamp assigned_at in the database rather than trusting the client to send an
-- honest clock. Not SECURITY DEFINER, so it carries no privilege of its own and
-- needs no revoke from PUBLIC (calling it outside a trigger just errors).
create or replace function public.feedback_log_stamp_assignment()
returns trigger
language plpgsql
as $$
begin
  if tg_op = 'INSERT' then
    if new.assigned_to is not null then
      new.assigned_at := now();
    end if;
  elsif new.assigned_to is distinct from old.assigned_to then
    -- Unassigning clears the stamp; reassigning restarts it.
    new.assigned_at := case when new.assigned_to is null then null else now() end;
  end if;
  return new;
end
$$;

drop trigger if exists feedback_log_stamp_assignment on public.feedback_log;
create trigger feedback_log_stamp_assignment
  before insert or update on public.feedback_log
  for each row execute function public.feedback_log_stamp_assignment();

commit;
