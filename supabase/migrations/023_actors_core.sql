-- =============================================================================
-- 023_actors_core.sql  —  Phase A: the Actor model (additive, reversible)
-- Spec: ComeWith_Data_Architecture_Concept.md §1.1, §6 Phase A.
--
-- Unifies people+orgs into ONE table; roles become relationships. Backfills from
-- artists/contractors/clients/sponsors with best-effort dedupe (Q1), PRESERVING
-- provenance so the relationship-inspection UI can verify/merge/fix by hand.
--
-- ADDITIVE ONLY: does NOT drop or convert artists/contractors/clients/sponsors
-- (the dashboard still WRITES to them; converting to views now would break writes).
-- The view-cutover happens later, after the dashboard is repointed to actors.
-- See BUILD_LOG_data_architecture.md §3 Phase A for the rationale/deviation.
--
-- RLS: admin-only (is_admin()) on all new tables. NO actor-self policies yet and
-- NO actor logins provisioned — that is gated to Phase C2 behind the financial-view
-- lockdown + negative tests (BUILD_LOG §2). NEVER blanket-grant anon (013/016/017/019).
-- NOT APPLIED by this commit — review before apply (dedupe especially).
-- =============================================================================

-- ----------------------------------------------------------------------------
-- Tables
-- ----------------------------------------------------------------------------
create table public.actors (
  id            uuid primary key default gen_random_uuid(),
  kind          text not null default 'person' check (kind in ('person', 'org')),
  display_name  text not null,
  legal_name    text,
  email         text,
  phone         text,
  instagram     text,
  website       text,
  notes         text,
  user_id       uuid references public.profiles(id) on delete set null, -- login link (nullable; no login enabled yet)
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now(),
  deleted_at    timestamptz
);
create index idx_actors_display_name on public.actors(lower(display_name)) where deleted_at is null;
create index idx_actors_email on public.actors(lower(email)) where deleted_at is null and email is not null;
create index idx_actors_user_id on public.actors(user_id) where user_id is not null;

create trigger set_updated_at before update on public.actors
  for each row execute function public.handle_updated_at();

alter table public.actors enable row level security;
create policy "Admins can manage actors" on public.actors for all using (public.is_admin());

create table public.actor_roles (
  id          uuid primary key default gen_random_uuid(),
  actor_id    uuid not null references public.actors(id) on delete cascade,
  role        text not null check (role in (
                'artist','dj','contractor','customer','sponsor','team',
                'performer','painter','dancer','vendor','venue_contact','host','crew')),
  context     text,
  active      boolean not null default true,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);
create unique index idx_actor_roles_unique on public.actor_roles(actor_id, role);
create index idx_actor_roles_role on public.actor_roles(role);

create trigger set_updated_at before update on public.actor_roles
  for each row execute function public.handle_updated_at();

alter table public.actor_roles enable row level security;
create policy "Admins can manage actor roles" on public.actor_roles for all using (public.is_admin());

-- Provenance: which legacy row(s) each actor was merged from. Powers the
-- inspection UI (show/merge/split) and makes the dedupe reversible.
create table public.actor_source_links (
  id            uuid primary key default gen_random_uuid(),
  actor_id      uuid not null references public.actors(id) on delete cascade,
  source_table  text not null check (source_table in ('artist','contractor','client','sponsor')),
  source_id     uuid not null,
  created_at    timestamptz not null default now()
);
create unique index idx_actor_source_links_unique on public.actor_source_links(source_table, source_id);
create index idx_actor_source_links_actor on public.actor_source_links(actor_id);

alter table public.actor_source_links enable row level security;
create policy "Admins can manage actor source links" on public.actor_source_links for all using (public.is_admin());

-- ----------------------------------------------------------------------------
-- Backfill (best-effort dedupe by email-then-name; provenance + roles preserved)
-- ----------------------------------------------------------------------------
-- Source priority for choosing the surviving actor's attributes when a match
-- group spans tables: client (1) > sponsor (2) > artist (3) > contractor (4).
create temp table _bf on commit drop as
  select 'client'::text src_table, c.id src_id, 1 prio, 'person'::text kind,
         c.full_name display_name, null::text legal_name, c.email, c.phone,
         null::text instagram, c.company website, c.user_id, 'customer'::text role
    from public.clients c where c.deleted_at is null
  union all
  select 'sponsor', s.id, 2, 'org',
         s.name, null, s.contact_email, s.contact_phone,
         null, s.website, null, 'sponsor'
    from public.sponsors s where s.deleted_at is null
  union all
  select 'artist', a.id, 3, 'person',
         a.stage_name, a.legal_name, a.contact_email, a.contact_phone,
         a.social_links->>'instagram', null, null, 'artist'
    from public.artists a where a.deleted_at is null
  union all
  select 'contractor', ct.id, 4, 'person',
         coalesce(ct.stage_name, ct.full_name), ct.full_name, ct.email, ct.phone,
         null, null, null, 'contractor'
    from public.contractors ct where ct.deleted_at is null;

alter table _bf add column match_key text;
update _bf set match_key = coalesce(nullif(lower(trim(email)), ''), lower(trim(display_name)));

-- 1) one actor per match_key, attributes from the highest-priority source row
create temp table _akeys on commit drop as select match_key, null::uuid actor_id from _bf where false;
with chosen as (
  select distinct on (match_key) match_key, kind, display_name, legal_name,
         email, phone, instagram, website, user_id
    from _bf
   order by match_key, prio
), ins as (
  insert into public.actors (kind, display_name, legal_name, email, phone, instagram, website, user_id)
  select kind, display_name, legal_name, email, phone, instagram, website, user_id from chosen
  returning id, coalesce(nullif(lower(trim(email)), ''), lower(trim(display_name))) as match_key
)
insert into _akeys (match_key, actor_id) select match_key, id from ins;

-- 2) provenance: link every legacy row to its actor
insert into public.actor_source_links (actor_id, source_table, source_id)
select k.actor_id, b.src_table, b.src_id
  from _bf b join _akeys k using (match_key)
on conflict (source_table, source_id) do nothing;

-- 3) roles: one row per (actor, role) the legacy data implies
insert into public.actor_roles (actor_id, role)
select distinct k.actor_id, b.role
  from _bf b join _akeys k using (match_key)
on conflict (actor_id, role) do nothing;

-- ----------------------------------------------------------------------------
-- Repoint sponsorships at actors (keep sponsor_id during transition)
-- ----------------------------------------------------------------------------
alter table public.sponsorships add column actor_id uuid references public.actors(id);
create index idx_sponsorships_actor_id on public.sponsorships(actor_id);

update public.sponsorships sp
   set actor_id = sl.actor_id
  from public.actor_source_links sl
 where sl.source_table = 'sponsor' and sl.source_id = sp.sponsor_id;

-- Grants: rely on 013 ALTER DEFAULT PRIVILEGES (new tables already covered). Do
-- NOT add anon grants. These tables expose no data to anon/authenticated because
-- their only policy is is_admin() (default-deny for everyone else).

-- =============================================================================
-- DOWN (manual reverse if needed, additive so low-risk):
--   alter table public.sponsorships drop column actor_id;
--   drop table public.actor_source_links;
--   drop table public.actor_roles;
--   drop table public.actors;
-- (Legacy artists/contractors/clients/sponsors are untouched — no data loss.)
-- =============================================================================
