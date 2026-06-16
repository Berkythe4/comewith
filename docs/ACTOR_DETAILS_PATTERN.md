# Actor Details Pattern

> Established in the **Event Hub sprint** (migration `031_v_actor_full.sql`). This is the
> convention the **Artist** and **Vendor** module sprints follow. Read this before adding any
> role-specific fields to a person or org.

## The rule

**`public.actors` is the universal core.** Every person and every org is one `actors` row
(`kind in ('person','org')`). Roles are **relationships** in `public.actor_roles`
(`artist, dj, contractor, customer, sponsor, team, performer, painter, dancer, vendor,
venue_contact, host, crew, donor`), not columns. One actor can hold many roles at once.

**Role-specific fields go in a one-to-one, nullable `actor_<role>_details` table** — never as
wide sparse columns on `actors`, and never as a JSON blob for anything you'll filter, sort, or
aggregate on.

```
actors (core: display_name, kind, email, phone, instagram, website, user_id, …)
  │
  ├─ actor_roles            (actor_id, role)            -- which hats this actor wears
  ├─ actor_artist_details   (actor_id PK→actors)        -- Artist sprint: genres, rate, rider…
  └─ actor_vendor_details   (actor_id PK→actors)        -- Vendor sprint: service_type, coi…
```

Each details table:
- has `actor_id uuid primary key references public.actors(id) on delete cascade` (true 1:1),
- holds only fields meaningful for that role,
- is **nullable / optional** — an actor with the `artist` role need not have a details row yet,
- carries its own `created_at/updated_at` + `set_updated_at` trigger,
- is RLS `for all using (public.is_admin())` until the external-access phase deliberately opens it.

## `v_actor_full` — the canonical read

`v_actor_full` is the one place code asks "give me an actor with their roles":

```sql
create or replace view public.v_actor_full
with (security_invoker = true) as
select
  a.*,
  coalesce((select array_agg(ar.role order by ar.role)
              from public.actor_roles ar
             where ar.actor_id = a.id and ar.active = true), '{}') as roles
from public.actors a
where a.deleted_at is null;
```

**As each `actor_<role>_details` table is added, extend this view with a `LEFT JOIN`** so the
new fields ride along on the same read (left join, so actors without that role/details still
return). Example shape the Artist sprint will produce:

```sql
create or replace view public.v_actor_full
with (security_invoker = true) as
select a.*, <roles array>,
       aad.genres, aad.rate, aad.rate_unit,      -- from actor_artist_details
       avd.service_type, avd.coi_on_file         -- from actor_vendor_details
from public.actors a
left join public.actor_artist_details aad on aad.actor_id = a.id
left join public.actor_vendor_details avd on avd.actor_id = a.id
where a.deleted_at is null;
```

## Security convention (load-bearing)

Create the view (and the extended versions) **`with (security_invoker = true)`**. A default
Postgres view runs with its owner's rights and **bypasses the underlying tables' RLS** — the
exact definer-bypass that forced the financial views to be `anon`-revoked. `security_invoker`
makes the view enforce the underlying `actors` / `actor_roles` / `actor_*_details` RLS against
*each caller*: admins pass `is_admin()`, any future external authenticated actor gets zero
rows, and `anon` is revoked outright. This keeps the convention safe to grant to
`authenticated` even after external logins arrive. **Never blanket-grant `anon`** (the
013/016/017/019 discipline).

## What NOT to do

- ❌ Add `artist_genres`, `vendor_service_type`, … columns directly onto `actors`.
- ❌ Stuff analyzable fields into an `actors.details jsonb` blob.
- ❌ Re-introduce the legacy `artists` / `sponsors` / `clients` tables as the home for new
  role data — those are being retired; the actor model is the source of truth.
- ❌ Create a details view as a plain (definer) view — always `security_invoker = true`.
