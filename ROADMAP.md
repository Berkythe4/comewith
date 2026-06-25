# Come With — Platform Roadmap

This document is the single source of truth for the Supabase migration plan
and the target architecture. Status colors and arrows are current as of
**2026-05-28** (Phase 2 close).

The phase numbering below is a **redesign** of the original Claude-generated
plan. Original used 0-12 with several phases (4, 6, 7, 10, 11, 12) never
specified; "frontend rewrites" was lumped into a single Phase 3 even though
the work spans admin tools, public pages, and an Edge Function backend.
This version uses 12 phases (0-11) with each phase being independently
shippable on staging.

**Status as of 2026-05-28 close of Phase 11**: ALL PHASES DONE. Migration
complete. comewith.org is live on Supabase prod. Known issues + redesign
items collected in project-phase-11-status memory.

---

> ## ⛔ PRIORITY CONTEXT (read first)
> **Come With is MAINTENANCE-ONLY.** The **CWF (Come With Fitness) BRD** is project #1 —
> **due June 15, 2026**. Nothing Come With Fitness goes in *this* repo
> (dashboard / schema / pages) until the BRD is done **and** there's an explicit go
> (LEARNINGS §5). Everything below is the Come-With dev roadmap, reconciled **2026-06-02**.

---

## Phase roadmap

```mermaid
flowchart TD
    classDef done fill:#22c55e,stroke:#166534,color:#fff
    classDef next fill:#eab308,stroke:#854d0e,color:#000
    classDef planned fill:#94a3b8,stroke:#1e293b,color:#fff

    P0["<b>Phase 0</b><br/>Schema + RLS + Storage<br/><i>migrations 001-012</i>"]
    P1["<b>Phase 1</b><br/>Data migration<br/><i>88 rows from Sheets</i>"]
    P2["<b>Phase 2</b><br/>Auth bootstrap<br/><i>magic-link + RLS proven</i>"]
    P3["<b>Phase 3</b><br/>Admin dashboard — read<br/><i>dashboard.html ↔ Supabase</i>"]
    P4["<b>Phase 4</b><br/>Admin dashboard — write<br/><i>agreements, income, equipment</i>"]
    P5["<b>Phase 5</b><br/>Edge Functions: transactional<br/><i>send-agreement, inquiry-notify</i>"]
    P6["<b>Phase 6</b><br/>Customer-facing flows<br/><i>inquiry form, customer portal</i>"]
    P7["<b>Phase 7</b><br/>Dance Infusion event hub<br/><i>event pages, ticketing import</i>"]
    P8["<b>Phase 8</b><br/>Mailing list<br/><i>subscribe, confirm, unsubscribe</i>"]
    P9["<b>Phase 9</b><br/>Resend broadcasts + webhooks<br/><i>Audiences sync, delivery events</i>"]
    P10["<b>Phase 10</b><br/>pg_cron automation<br/><i>nightly mv refresh, scheduled sends</i>"]
    P11["<b>Phase 11</b><br/>Production cutover + hardening<br/><i>DNS, sunset Apps Script, monitoring</i>"]

    P0 --> P1 --> P2 --> P3 --> P4 --> P5
    P5 --> P6
    P5 --> P7
    P6 --> P8
    P8 --> P9
    P4 --> P10
    P9 --> P11
    P10 --> P11
    P7 --> P11

    class P0,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11 done
```

### Phase descriptions and dependencies

| # | Name | What it produces | Depends on |
|---|---|---|---|
| 0 ✅ | Schema + RLS + Storage | 30 tables, helper fns, audit triggers, 6 buckets, RLS policies | — |
| 1 ✅ | Data migration | 88 rows imported from Google Sheets (clients, contractors, equipment, expenses, income) | 0 |
| 2 ✅ | Auth bootstrap | magic-link configured, Berky=master_admin, RLS isolation proven | 0 |
| 3 ✅ | Admin dashboard — read | `dashboard-v2.html` reads from Supabase (all 7 admin tables: inquiries, agreements, clients, income, expenses, equipment, events). Magic-link login. No writes yet — read-only de-risks the schema before exposing edits | 2 |
| 4 ✅ | Admin dashboard — write | Inquiry status writes, agreement status writes, "Add income" modal, "Add expense" modal with receipt upload to Storage. Daily ledger work runnable from the dashboard. Full CRUD on clients/equipment/events deferred to 4.5 or Phase 7 | 3 |
| 5 ✅ | Edge Functions: transactional | `send-agreement` (creates token, emails sign link via Resend, marks agreement sent), `get-agreement-by-token` (public, returns agreement for signing page), `mark-signed` (records typed-name signature, notifies all master_admins). `sign.html` customer signing page. Web-based signing — no PDFs. `inquiry-notify` deferred to Phase 6 alongside the public form | 4 |
| 6 ✅ | Customer-facing flows | `index-v2.html` public inquiry form (anon insert via `return=minimal` workaround), `customer_portal-v2.html` shows signed-in customer's agreements with "Review & sign" deep link, `inquiry-notify` Edge Function emails master_admins on new submission. **Anon-RLS resolved** — was a `Prefer: return=representation` quirk, not a policy bug | 5 |
| 7 ✅ | Dance Infusion event hub | DI2 seeded (1 venue, 1 event, 9 sponsors+sponsorships, 5 artists+bookings, 5 raffle prizes, 4 expenses, 16 RA tickets). `events/dance-infusion-2/index-v2.html` public hub via `get-event-hub` Edge Function. Dashboard-v2 admin tabs for sponsors/sponsorships/artists. `import_ticketing.py` CSV importer (RA + extensible for Zeffy) | 5 |
| 8 | Mailing list | Public subscribe form, double-opt-in confirmation, unsubscribe via tokenized URL, segments. Self-hosted per decision #8 | 6 |
| 9 | Resend broadcasts + webhooks | Audiences sync, campaign drafting UI, send queue, webhook handler for `delivered`/`bounced`/`complained` → `mailing_events` table | 8 |
| 10 | pg_cron automation | Nightly materialized view refresh, scheduled mailing sends, audit log retention, weekly digest emails. Lives in `automation_jobs` table; pg_cron calls Edge Functions via pg_net | 4 |
| 11 | Production cutover + hardening | Run migrations on `comewith-prod`, DNS swap if needed, Netlify env vars updated, Apps Script triggers disabled and code archived, monitoring dashboards, backup verification | 7, 9, 10 |

**Why this re-decomposition vs. the original:**

- Original Phase 3 = "frontend rewrites + real-time admin dash" tried to fit
  every HTML page rewrite into one phase. In practice that's weeks of work.
  Splitting into admin-read (P3), admin-write (P4), customer-facing (P6), and
  event hub (P7) lets each ship and prove itself independently.
- Edge Functions weren't in the original plan as their own phase — they were
  implicit in "Phase 9: pg_cron + Edge Functions". But transactional functions
  (send agreement, inquiry notification) need to exist before the customer-
  facing pages can rely on them (P5 → P6).
- Original Phase 5 = "production cutover" was placed early, but it can't
  meaningfully happen until P10 (automation) is verified on staging. Renamed
  P11 and pushed to the end where it belongs.
- Phases 4, 6, 7, 10, 11, 12 in the original were never specified. This
  proposal fills them.

---

## End-state architecture

This is the target state after Phase 11. Apps Script + Google Sheets do not
appear — they're sunset by then.

```mermaid
flowchart LR
    classDef user fill:#dbeafe,stroke:#1d4ed8,color:#000
    classDef page fill:#fef3c7,stroke:#a16207,color:#000
    classDef sb fill:#86efac,stroke:#166534,color:#000
    classDef ext fill:#fbcfe8,stroke:#9d174d,color:#000

    subgraph USERS["👥 Users"]
        direction TB
        VIS[Visitor]
        CUST[Signed-in customer]
        ADMIN[Berky — master_admin]
    end

    subgraph NETLIFY["🌐 Netlify (static hosting at comewith.org)"]
        direction TB
        IDX[index.html<br/>marketing + inquiry form]
        EQP[equipment_list.html<br/>public rental catalog]
        EVT["/events/dance-infusion-*<br/>event hub pages"]
        AUTH_PG[auth.html<br/>magic-link entry]
        PORT[customer_portal.html<br/>my agreements + PDFs]
        DASH[dashboard.html<br/>admin workspace]
    end

    subgraph SB["🟩 Supabase platform"]
        direction TB
        SBA[Auth<br/>magic-link, JWT]
        REST[PostgREST API<br/>CRUD with RLS enforced]
        STG[Storage<br/>6 buckets: agreements,<br/>photos, receipts, logos]
        EDGE[Edge Functions<br/>send-agreement, inquiry-notify,<br/>mailing-confirm, resend-webhook,<br/>mailing-send, mv-refresh]
        CRON[pg_cron<br/>scheduled triggers]
        DB[(Postgres<br/>30 tables + views<br/>RLS policies<br/>audit_log)]
    end

    subgraph EXT["📨 External services"]
        direction TB
        RES[Resend<br/>transactional + broadcast]
        DNS[Namecheap DNS<br/>SPF, DKIM, MX]
        ZEF[Zeffy / Resident Advisor<br/>ticketing CSVs]
    end

    VIS --> IDX
    VIS --> EQP
    VIS --> EVT
    VIS --> AUTH_PG
    CUST --> PORT
    CUST --> AUTH_PG
    ADMIN --> DASH
    ADMIN --> AUTH_PG

    IDX -->|insert inquiry| REST
    EQP -->|read available| REST
    EVT -->|read event,<br/>sponsors, artists| REST
    AUTH_PG -->|sign in| SBA
    PORT -->|read own agreements| REST
    PORT -->|signed URLs| STG
    DASH -->|CRUD all tables| REST
    DASH -->|upload receipts,<br/>signed PDFs| STG
    DASH -->|trigger sends| EDGE

    REST --> DB
    EDGE -->|service-role| DB
    EDGE -->|send| RES
    EDGE -->|sign URLs| STG
    SBA -.->|JWT to anon/<br/>authenticated| REST
    CRON -.->|via pg_net| EDGE
    CRON -->|refresh mv| DB
    RES -.->|delivery webhooks| EDGE
    DNS -.->|verifies| RES
    ZEF -.->|periodic import| EDGE

    class VIS,CUST,ADMIN user
    class IDX,EQP,EVT,AUTH_PG,PORT,DASH page
    class SBA,REST,STG,EDGE,CRON,DB sb
    class RES,DNS,ZEF ext
```

### Request-flow examples

| Scenario | Path through the system |
|---|---|
| New website inquiry | `index.html` → PostgREST → `inquiries` table (anon INSERT policy) → Resend trigger via `inquiry-notify` Edge Function → Berky's email |
| Customer signs agreement | Berky drafts in `dashboard.html` → PostgREST → `agreements` (admin policy) → `send-agreement` Edge Function generates PDF → uploads to `agreements` Storage bucket → Resend sends signing link → customer clicks tokenized `agreement_links` URL |
| Customer logs in | `auth.html` magic-link form → Supabase Auth sends email → click sets JWT → `customer_portal.html` reads JWT, calls PostgREST with `auth.uid()` → RLS filters `agreements` to their own client only |
| Nightly KPI refresh | `pg_cron` job fires at 03:00 → `pg_net` calls `mv-refresh` Edge Function → function executes `refresh materialized view concurrently mv_cross_event_kpis` (and the other two MVs) |
| Bounce handling | Resend webhook → `resend-webhook` Edge Function (verifies signing secret) → inserts row in `mailing_events` → if event is `bounced`/`complained`, updates `subscribers.status` |

---

## How to read status colors

- 🟢 **green** — phase complete, validated on staging
- 🟡 **yellow** — phase is the next one up, not started
- ⚪ **gray** — planned, dependencies not yet met

Update the `class P# done|next|planned` lines in the Mermaid block as phases
progress. The mermaid diagrams render natively on GitHub and in VSCode/Cursor
preview — no build step required.

---

## Current state — reconciled 2026-06-24 (Operations + CRM build-out)

> **Priorities/sequencing owned in the planning chat; this file = dashboard execution backlog.**

### ✅ Done / LIVE on prod — operations + CRM build-out (2026-06-24)
Migrations **045–051 all APPLIED to prod**; `dashboard.html` pushed (Netlify live); `send-campaign` redeployed.

**Actor model is now the single source of truth.**
- **047 — legacy person/org tables RETIRED**: dropped `clients`, `sponsors`, `artists`, `contractors`,
  `artist_bookings`, `artist_notes` (+ orphaned `v_sponsor_history`/`v_artist_history`/`mv_repeat_sponsors`/
  `mv_top_artists`; trimmed the MV refresh cron). `inquiries`/`agreements`/`income`/`mileage` actorized
  (`actor_id`; customer-self RLS repointed). Entire dashboard reads `actors` — Sponsors/Sponsorships/
  Clients/Artists tabs + every picker. **No legacy-table references remain.**
- **Actors management tab** (048 `actors.status`; 049 `actors.org_id` + widened role vocab): one
  inline-editable, sortable table — add / on-hold / archive, role chips, multi-select role + kind filters,
  person→org affiliation, org→venue assignment. (Closes the "Sponsor/Artist/Vendor tab repoint" backlog items.)

**Event hub.**
- **Files tab** (045 `document_types` + `files.vendor_actor_id`): replaced Contracts; doc-type buckets
  (+ add custom), per-bucket upload, vendor = vendor-role actors only, surfaces contract-attached files
  (recovered the "lost" Signal contract).
- **Customers tab**: deduped union of participants/attendees/donors/sponsors with ticket counts + reconciliation block.
- **Overview**: type-aware (Come With Production = service framing); **Engagement & marketing** section
  (per-event IG snapshots back-dated via Log-IG; marketing spend); checklist moved to bottom.
- **Equipment** (046 `equipment_components` + `wishlist` status): inventory reconciled to the Financial
  master — serials, 6 buckets (DJ/Sound/Camera/Misc/Wire/Accessories), prices; fixed 3 wrong "retired";
  edit modal + usage history; **gear bundling** (parent↔child, auto-included on event assignment);
  wishlist; category filter on the assign popup. (Closes "Full Equipment module".)
- Venue event-history; richer Edit-core (venue **auto-fills capacity**, doors/end/description);
  new series **Come With Production** + **Content Creation** (+ `seriesToType()` keeps `events.type` in sync).

**Money / KPIs.**
- **DI#1 + DI#2 fully itemized & tied to the audit** (tickets→`ticketing`, donors→`third_party_donations`,
  per-buyer→`guest_event_attendance`); historical events backfilled (Maxwell→production, showcases→content,
  DJs incl. SPF 50). DI#1 net = $1,140-to-MS exactly.
- **Events page rebuilt as the command center**: aggregates in **three separate money models**
  (Come With Parties / Production & content ≈net-$0 / Dance Infusion charity), type-adaptive per-event
  "Result", filters (series/status/year/search/**completed-only**), split marketing cards.
- **Expenses → events** (050 `event_na`): inline event dropdown, N/A-overhead flag (Software/Equipment
  auto-N/A), clean "needs an event" list; **Simplifi import** (Category='Work Expenses', deduped, 63 added,
  5 auto-assigned, tagged `[simplifi 0624]`).
- **Strategy KPI fix** (051 `v_kpi_computed`): event-derived cards (di.*/parties.*) compute **LIVE**
  (completed-only) instead of null; +4 cards (To MS total, Net P&L total, Mailing list, Repeat attendees);
  "last updated" on cards; **audience + attendance trend sparklines**.

**Email / campaigns** (stack already deployed): fresh `RESEND_API_KEY` set; `send-campaign` gains
**test-send** + per-campaign **stats**; Campaigns tab segment picker + preview + audience confirm.
**External TODO before first blast: verify comewith.org as a sending domain in Resend.**

**Open / offered (not built):** YouTube/Instagram auto-pull API (need YT API key + channel ID / Meta app +
IG Business token); in-app Simplifi importer; ~12 small "needs an event" expenses (Elements) for Keith to N/A;
optional sortable Events headers + follower-growth-vs-prior-event delta.

### ✅ Done / LIVE on prod — staff access + social calendar (2026-06-23)
- **Migration 041 — staff access model** (APPLIED to prod): data-driven, grouped nav
  (**Sales / Operations / Finance / Partners / Audience / Insights**) rendered from
  `module_registry`; `profiles.staff_role` (`operations` / `marketing` / `full`);
  `module_registry` + `user_module_access` (+ RLS); `user_can_access_module()` helper.
  Client-side **signed-off badge gate** (non-master staff see only modules that are
  signed off **and** in their role scope, with per-user grant/revoke overrides) and a
  master-only **Team tab** (set roles, per-user access, and flip module sign-off).
  Existing sub_admin (liz@comewith.org) backfilled to `staff_role='full'`.
  **Signed off so far: Events, Team, Social Calendar** (everything else built-not-released).
- **Migration 044 — Social Calendar** (APPLIED to prod): `social_posts` + `social_post_notes`
  with **real per-module RLS** (clean leaf tables, no Events-hub coupling, so RLS was safe to
  apply here). Stage pipeline idea→drafted→review→planned→scheduled→posted→archived;
  **Timeline (default) / Board / List** views; full post CRUD; **threaded, timestamped notes**;
  read-only **snapshot export** (self-contained chronological timeline → Print / Save-as-PDF)
  to share with collaborators (e.g. Janelle) **without a login**.
- **Deployed:** `staff-access-model` merged to master (`8fa2a62`), pushed; Netlify serving the
  new `dashboard.html` to comewith.org. DB migrations **041 + 044 applied to prod**; the nav
  also added a Social Calendar module and a master-only Team module.

### 🔶 STAGED — committed but NOT applied (next deliberate session, in this order)
Files are on master (`042`/`043` SQL, `invite-user/index.ts`) but inert — Netlify doesn't run
them. Do these as one focused, rested session:
1. **Apply 042** (hard per-module RLS) — first resolve the 4 `[VERIFY]` policy-name markers
   (`task_templates`, `subscribers`, `mailing_campaigns`, `feedback_log`); the Events-hub
   dependency carve (`can_use_events_module()`) is already built. **Then apply 043**
   (`events.audited` publish gate: rebuilds the **5 money views** as `security_invoker` and
   CASE-gates the money columns on `is_master_admin() OR audited`; re-asserts anon-revoke on
   every rebuilt view). **Smoke-test with a throwaway staff account before trusting either.**
2. **Deploy `invite-user` Edge Function** (authored, in repo, not deployed) — enables the Team
   "＋ Add person" button:
   `SUPABASE_ACCESS_TOKEN=$SBP_PAT supabase functions deploy invite-user --project-ref yaytdosxfhcqatmhctzk`
3. **Create staff logins** — Martin (operations), Henry (operations), Janelle (marketing).
   **ONLY after 043** so they never have a window of ungated financial access.

**KNOWN GAP until 043 is applied:** any authenticated non-master (today only liz@comewith.org —
a `sub_admin`, `full` scope, has a password but **never signed in**) can read the financial views
via direct REST. The new nav only **hides** Finance; it does **not** RLS-gate it. No staff logins
exist yet, so not currently exploitable — but **043 must precede any staff login creation.**

### 🔁 Carrying forward
- **Round-2 module sign-off (ongoing process)** — as Keith reviews each module, flip it
  `signed_off` from the Team tab to release it to the staff roles whose scope includes it.
  Only **Events / Team / Social Calendar** are released today; the rest are built but gated.
- Events Services Agreement bug fixes + dashboard filter/sort (tracked under "Still-open bugs");
  MSA e-signatures (tracked under "Parked" — still hard-blocked by the financial-view fix below).

### ✅ Done / LIVE on prod — cumulative (through 2026-06-16)
- **Migration & cutover (0–11) + KPI/metrics + money model (015–022)** — Strategy tab, entry forms
  (Log Event / Numbers / Edit Target + create/retire-metric), feedback log + Notes, event edit +
  soft-delete, Add Sponsor, per-event Money panel; canonical revenue/P&L (net P&L incl. tickets).
- **DI #2 impact report + public audit + reusable `/staging/` gate** — "% to mission" framing,
  founder-contribution note. Gated; **not yet public** (consent sweep — see Parked).
- **Data architecture 023–028 APPLIED to prod AND POPULATED with reconciled DI data** — actors /
  roles / event_participants, content_items + tags, workflow (tasks / templates / contracts / files /
  budget_lines + variance / touchpoints), metric_definitions + v_data_points (+ nightly rollup).
  DI#1 loaded at **39% to mission**, DI#2 at **31%**; role-overlap proven (Keith = dj+donor+sponsor+
  team; Crossroads = vendor+sponsor); 17 actors, 5 DI#2 DJ participants, 12 sponsorships; no dup
  actors; anon-401 holds. (`events/dance-infusion/DI_DATA_LOAD_LOG.md`.)
- **Migration 029** — `sponsorships.sponsor_id` nullable (sponsorships attach to actors) + `actor_roles` `donor` role.
- **Tools deployed + admin-gated** on comewith.org — `/tools/actor-inspector|test-checklist|visualizer.html`.
- **MODULE SERIES — Event Hub & Guest layer (migrations 030–040), all live on master/Netlify:**
  - **Event Hub** (sprints 1–2): per-event detail page — Overview/People/Tasks/Money/Equipment/Contracts/Files,
    stage stepper, day-of generator, multi-role participants (`roles[]`), audit triggers, `v_actor_full`,
    bulk add/edit, inline Money fix, contract docs, IG-followers KPI capture.
  - **Venue / contact matrix** (sprint 3a): Venues tab, `venue_contacts`/`vendor_contacts`, "last time" lookup.
  - **Conditional workflows + template editor** (sprint 3b): `events.cw_providing_gear`, gear-aware generation,
    outreach templates auto-assigned via the matrix, in-dashboard Templates editor (future-only), grouped assign picker.
  - **Chain fix** (sprint 4): venue-save display fix, ONE gear task + equipment sheet, venue-as-counterparty
    (`venues.actor_id`), delete-stays-deleted; actors-only historical backfill (DI#1/showcase participants, Signal contacts).
  - **Attendee + Guest module** (sprints 5–7): `guests`/`subscribers`/`guest_event_attendance`, DI#2 ledger import
    (people-only, money never written), guest→actor graduation (`guests.actor_id`), Guests tab w/ lifetime stats +
    filters (subscribed = mailing list), returning-attendee KPI **fixed (fuzzy name+email; DI#2 returning 1→12)**,
    mission/business spend split (DI = MS-Society mission vs CW Parties business).
  - **Counts (prod):** guests **97**, subscribers **86** subscribed (11 opt-out-respected), actors **38**, attendance **108**.
  - **Money discipline held:** no financial rows written from any backfill; DI#1/DI#2 `v_event_summary` unchanged throughout.

### 🟡 Held (committed locally, push held for Keith)
`261797d` (029 + DI load log), `5cbb51e` (roadmap backlog). (Module-series work 030–040 is pushed to master.)

### 🚫 GATED BLOCKER (hard dependency — keep)
**Financial-view security fix — BEFORE any customer/external login:** revoke the 5 financial views
(+ `v_budget_variance`, `v_data_points`, `mv_event_data_points`) from `authenticated`, re-issue as
`security_invoker` over RLS-gated tables (non-admins get ZERO rows); negative tests pass on staging
first (`tools/test-checklist.html` → Security 🔴). ⚠ Covers **existing `customer`-role logins too**
(they're `authenticated`; views revoked from `anon` only today). The dormant actor-self RLS tier is
built (024/026) but **no non-admin login is provisioned**.
>
> **⚠ Update 2026-06-23 — staged migration 043 is the START of this fix, not the whole thing:**
> 043 rebuilds the **5 money views** (`v_event_summary`, `v_kpi_event_financials`, `v_kpi_parties`,
> `v_kpi_dance_infusion`, `v_kpi_dashboard`) as `security_invoker` and gates them on
> `is_master_admin() OR events.audited`. It does **NOT** yet touch **`v_budget_variance`,
> `v_data_points`, or `mv_event_data_points`** — extend 043 (or add a follow-up migration) to cover
> those three before declaring this blocker closed. Also note the **model differs**: 043 implements
> *staff-audited gating* (staff see audited events), whereas this blocker was written for
> *customers/external get zero rows*; reconcile the two intents when applying. As of now **one
> `sub_admin` login (liz) exists** and can still read all of these views via REST — see the staged
> "KNOWN GAP" above.

### 🟦 Queued (dashboard backlog — order owned in planning chat)
1. **Audit cleanup follow-up** — approve the 6 same-human guest↔actor links + review the 8 variant pairs
   in `GUEST_ACTOR_AUDIT.md` (Keith's manual call; includes family records like Francis/Theresa Berkman).
2. ~~Artist module~~ / ~~Vendor module~~ / ~~Sponsor tab repoint~~ / ~~Full Equipment module~~ — **DONE 2026-06-24**:
   legacy tables retired (047), unified **Actors** tab (roles incl. artist/vendor/sponsor + org/venue links),
   and the full equipment module (buckets, edit, usage history, bundling, wishlist) all shipped.
   Remaining slivers: per-role detail tables (`actor_artist_details`/`actor_vendor_details`) if richer
   per-role fields are wanted; equipment **rental ROI / rental-vs-own-use** view.
- Carryover smaller items: actor-inspector "Events" section; Tools nav in the dashboard;
  DI#2 thank-you/survey send; `third_party_donations` actor FK (donations still text `donor_name`);
  YouTube/Instagram auto-pull API; in-app Simplifi importer.

### ❓ Decisions waiting on Keith (block nothing)
- **Cold subscriptions** — keep or drop the ~20 attendees who never ticked RA marketing opt-in.
- **81-name door list** (no emails) — import comps as guests or leave out.
- **Audit merges** — which of the `GUEST_ACTOR_AUDIT.md` same-human links / variant pairs to approve
  (3 flagged pairs are likely *different* people — do not merge).

### ⏸ Parked (each its own session)
- **Actor onboarding + MSA e-signing** — hard-blocked by the financial-view fix above.
- **Roadmap planning tool** — buy not build (Notion vs Trello); separate from this dev roadmap.
- **Flywheel redesign** — `ComeWith_Strategy_Dashboard.html`.
- **`equipment_usage` UI wiring** into the Log Event panel (schema ready in 024).
- **KPI views repoint `series` → `type`** (series exact-match contract kept for now).
- **Impact report → public** — consent sweep (sponsors/team/artists/raffle; Yankees-hats donor) +
  fill placeholders + remove the 2 `/staging/` guard lines; set dashboard `di.cost_to_raise` → 60%-to-mission.
- **Smaller dev items** — Expenses CSV import; per-event line-item editor; reactivate-metric UI;
  `FORM_DEFS` `created_by`-aware handler; DI dashboard `event_date` display fix.

### 🐞 Still-open bugs (re-verify when next in that flow)
- Events Services Agreement (payment-method/recording-rights buttons, fee-total auto-update, client
  auto-populate); dashboard tabs need filter/sort.
