-- =============================================================================
-- 047_retire_legacy_tables_actorize.sql
--
-- Finish the actor migration: repoint the last client-coupled tables to actors
-- and DROP the empty legacy person/org tables (clients, sponsors, artists,
-- contractors) + their orphaned helper tables/views.
--
-- SAFE: every affected column/table verified to hold 0 rows on prod before this
-- migration (clients/sponsors/artists/contractors empty; agreements/mileage/
-- inquiries empty; income has 6 rows but 0 with client_id; sponsorships.sponsor_id,
-- raffle_prizes.donor_sponsor_id, artist_bookings/artist_notes/photos all 0).
-- The 4 legacy views/matviews (v_sponsor_history, v_artist_history,
-- mv_repeat_sponsors, mv_top_artists) are unused by the dashboard. Wrapped in a
-- transaction, so any missed dependency rolls the whole thing back.
-- =============================================================================
begin;

-- A) actorize the client-coupled tables (add actor_id; the legacy client_id is dropped in C)
alter table public.inquiries  add column if not exists actor_id uuid references public.actors(id) on delete set null;
alter table public.agreements add column if not exists actor_id uuid references public.actors(id) on delete set null;
alter table public.income     add column if not exists actor_id uuid references public.actors(id) on delete set null;
alter table public.mileage    add column if not exists actor_id uuid references public.actors(id) on delete set null;
create index if not exists idx_inquiries_actor  on public.inquiries(actor_id)  where actor_id is not null;
create index if not exists idx_agreements_actor on public.agreements(actor_id) where actor_id is not null;
create index if not exists idx_income_actor     on public.income(actor_id)     where actor_id is not null;

-- A2) repoint the customer-self RLS policy from clients(client_id) to actors(actor_id)
--     (this is the only policy that referenced the legacy column).
drop policy if exists "Customers can read own agreements" on public.agreements;
create policy "Customers can read own agreements" on public.agreements
  for select using (actor_id in (select id from public.actors where user_id = auth.uid()));

-- B) drop the orphaned legacy views/matviews + stop the cron from refreshing the dropped MVs
drop materialized view if exists public.mv_repeat_sponsors;
drop materialized view if exists public.mv_top_artists;
drop view if exists public.v_sponsor_history;
drop view if exists public.v_artist_history;
select cron.schedule('refresh-materialized-views', '0 3 * * *',
  'refresh materialized view concurrently public.mv_cross_event_kpis;');

-- C) drop legacy coupling columns (all empty) so the parent tables can go
alter table public.inquiries     drop column if exists client_id;
alter table public.agreements    drop column if exists client_id;
alter table public.income        drop column if exists client_id;
alter table public.mileage       drop column if exists client_id;
alter table public.sponsorships  drop column if exists sponsor_id;
alter table public.raffle_prizes drop column if exists donor_sponsor_id;
alter table public.photos        drop column if exists artist_id;

-- D) drop the legacy tables (superseded by the actor model; all 0 rows)
drop table if exists public.artist_bookings;
drop table if exists public.artist_notes;
drop table if exists public.artists;
drop table if exists public.contractors;
drop table if exists public.sponsors;
drop table if exists public.clients;

commit;

-- DOWN: not reversible in practice (tables dropped). Restore from backup if needed.
