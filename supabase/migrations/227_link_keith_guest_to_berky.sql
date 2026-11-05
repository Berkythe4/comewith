-- =============================================================================
-- 227_link_keith_guest_to_berky.sql - one person, one roster row
-- =============================================================================
-- The Control Center and the event hub's Customers tab now read ONE roster
-- (CC.eventRoster), which joins lineup actors to door-list guests through
-- guests.actor_id. Keith's guest record ("Keith Berkman", the DJ comp on DI #3)
-- was never linked to his DJ actor "Berky", so the roster showed Berky and
-- Keith Berkman as two people. Link them. Nothing else changes; the record's
-- name and email stay as they are.
-- =============================================================================

begin;

do $$
declare n int;
begin
  if exists (select 1 from public.guests where actor_id = '09939421-2095-4151-b943-2d943bf83615' and deleted_at is null) then
    raise exception 'another guest is already linked to Berky';
  end if;
  update public.guests set actor_id = '09939421-2095-4151-b943-2d943bf83615'
   where id = 'bd39d389-a3a7-4102-bbca-5d0f4d7ba195' and full_name = 'Keith Berkman' and actor_id is null;
  get diagnostics n = row_count;
  if n <> 1 then raise exception 'Keith Berkman guest not found unlinked'; end if;
end $$;

commit;
