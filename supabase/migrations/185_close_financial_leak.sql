-- ============================================================
-- COME WITH — 185 SECURITY: the event financial gate never checked WHO
--
-- LIVE DATA LEAK, closed here. Anonymous callers could read the ledger:
--     expenses               29 rows  (dates, amounts, payee names, categories)
--     ticketing              59 rows
--     third_party_donations  16 rows  (donor names and amounts)
--     sponsorships           12 rows
--     income                  9 rows
-- Verified against prod with the real publishable key before and after.
--
-- THE BUG. 043 put this policy on all five tables:
--     for select using (can_see_event_financials(event_id))
-- and defined the helper as:
--     select public.is_master_admin()
--         or (p_event_id is not null
--             and exists (select 1 from events e
--                          where e.id = p_event_id and e.financials_released));
-- The second branch asks WHAT has been released and never WHO is asking. RLS
-- policies apply to role `public`, which includes `anon`, so the moment an event
-- had financials_released = true its money rows were readable by the entire
-- internet. The intent - "staff see an event's money only once it is released" -
-- was only ever half expressed: the release condition was written, the staff
-- condition was assumed.
--
-- THE FIX. Add the missing half. is_admin() is master_admin or sub_admin, and it
-- already returns false for anon (no profile for a null auth.uid()) and for a
-- deactivated profile (the 098 deleted_at contract). The 043 behaviour is
-- otherwise unchanged: master sees everything, staff see released events,
-- customers and listeners see nothing, anon sees nothing.
--
-- WHY THE VIEWS WERE FINE AND THE TABLES WERE NOT. Every financial VIEW is
-- anon-revoked (decision E1) and answers 401, which is what the post-apply check
-- and the REST spot-check both look at. The underlying TABLES carry an anon grant
-- from 013's default privileges and were relying on RLS alone - so they answered
-- 200, and the rows came with it. A grant check would never have caught this;
-- only reading the body does.
--
-- HOW IT WAS MISSED FOR SO LONG: every REST spot-check in this repo was run with
-- an empty apikey, because .env has no SUPABASE_ANON_KEY - the variable is
-- SUPABASE_PROD_PUBLISHABLE_KEY. An empty key answers 401 for everything, public
-- or not, so every check appeared to pass. See the note added to post_apply.sql.
-- ============================================================
begin;

create or replace function public.can_see_event_financials(p_event_id uuid)
returns boolean
language sql
stable
security definer
set search_path to 'public'
as $$
  select public.is_master_admin()
      or (public.is_admin()                 -- <- 185: the half that was missing
          and p_event_id is not null
          and exists (select 1 from public.events e
                       where e.id = p_event_id and e.financials_released));
$$;

comment on function public.can_see_event_financials(uuid) is
  'Can the CURRENT CALLER see this event''s money? Master always; other staff '
  'once the event''s financials are released. Both halves matter - checking only '
  'that the event was released let anon read the ledger (fixed in 185).';

-- A row with no event_id (company overhead, general income) was never covered by
-- the released branch and is master-only, which is correct and unchanged.

commit;

-- DOWN: restore the 043 definition. Do not - it is the leak.
