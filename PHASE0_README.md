# Come With — Phase 0 Implementation Pack

This folder is the complete Phase 0 deliverable: SQL migrations, RLS policies, storage bucket setup, and the environment-variable contract. Drop it into your `comewith` repo at the root.

## Folder structure after dropping this in

```
comewith/
├── index.html
├── dashboard.html
├── ...existing files...
├── .env.example                              ← new (template only — never commit real values)
└── supabase/                                 ← new folder
    └── migrations/
        ├── 001_extensions.sql
        ├── 002_profiles.sql
        ├── 003_clients_contractors.sql
        ├── 004_inquiries_agreements.sql
        ├── 005_financials.sql
        ├── 006_equipment.sql
        ├── 007_events.sql
        ├── 008_artists.sql
        ├── 009_mailing_list.sql
        ├── 010_automation_audit_photos.sql
        ├── 011_views.sql
        └── 012_storage.sql
```

## What this pack does NOT do

- It does not connect to your Supabase projects (no secrets are baked in)
- It does not modify any of your existing HTML files
- It does not run any migrations — you do that yourself in Claude Code (Phase 1)

## How to use it — order of operations

### Step A — drop these files into your repo

1. Open your `comewith` folder in File Explorer
2. Copy the entire `supabase` folder from this pack into your repo root
3. Copy `.env.example` into your repo root
4. In Command Prompt:
   ```
   cd Documents\comewith
   git add supabase .env.example
   git commit -m "Phase 0: add Supabase schema and env contract"
   git push origin master
   ```

### Step B — confirm .gitignore excludes secret files

Open the `.gitignore` file in your repo root (create one if it doesn't exist). It must contain at least these lines:

```
.env
.env.local
.env.production
.env.*.local
*.secret
```

If it doesn't have those lines, add them, then commit.

### Step C — run migrations on STAGING first (Phase 1 work, do this in your next Claude Code session)

In Claude Code, open Supabase staging in your browser at the same time:

1. Go to your `comewith-staging` project in the Supabase dashboard
2. Click **SQL Editor** in the sidebar
3. Run each migration file in order (001 → 012), pasting one file at a time:
   - Click **+ New query**
   - Paste the file contents
   - Click **Run**
   - Confirm "Success" before moving to the next file
4. After 012 finishes, check **Table Editor** in the sidebar — you should see all the tables listed

### Step D — sanity-check on STAGING

In Supabase SQL Editor, run these checks one by one:

```sql
-- Should return 20+ rows: every table you just created
select table_name from information_schema.tables
where table_schema = 'public' order by table_name;

-- Should return 12 rows: every migration
-- (Run only AFTER you set up migration tracking — for now just count tables)
select count(*) from information_schema.tables where table_schema = 'public';

-- Confirm RLS is enabled on every table
select tablename, rowsecurity from pg_tables
where schemaname = 'public' and rowsecurity = false;
-- ^ This should return ZERO rows. If any table is listed, RLS isn't enabled on it.

-- Confirm storage buckets exist
select id, public from storage.buckets order by id;
-- Should return 6 buckets: agreements, artist-photos, equipment-photos,
--                          event-photos, receipts, sponsor-logos
```

### Step E — repeat on PRODUCTION when staging is verified

Once staging is green, run the same 12 migrations on `comewith-prod`. This is Phase 5 work — don't do it until after the rest of the migration is tested.

## Reference: the 14 decisions baked into these files

| # | Decision | Where it shows up |
|---|---|---|
| 1 | Resend + SPF/DKIM | DNS records installed at Namecheap, sender identities in `.env.example` |
| 2 | Force-reset passwords | `must_change_password` flag in `profiles` table |
| 3 | PDFs to Supabase Storage | `agreements` bucket in 012, `signed_pdf_path` in agreements table |
| 4 | Real-time admin dash | Will be enabled only on admin queries during Phase 3 frontend rewrites |
| 5 | Magic-link customer auth | Configured in Supabase dashboard → Auth → Providers (Phase 2) |
| 6 | Separate staging project | Run all migrations on staging first, then production |
| 7 | Dance Infusion integrated Phase 1 | `events`, `venues`, `sponsors`, `sponsorships`, `guests`, `ticketing` tables all present |
| 8 | Self-hosted mailing list | `subscribers`, `subscriber_segments`, `mailing_campaigns`, `mailing_events` tables |
| 9 | pg_cron + Edge Functions | `pg_cron` extension enabled in 001; `automation_jobs` registry table in 010 |
| 10 | Artist directory admin-only | `artists` table RLS = `is_admin()` only |
| 11 | Storage transforms | Image buckets configured in 012 with public read |
| 12 | Free tier + Pro at 80% | Already on free tier; nothing to do in code |
| 13 | No Stripe | No payments tables in this pack |
| 14 | Backup repo | `comewith-archive-2026-05-pre-migration` (Step 1 of Phase 0, already done) |

## What comes next

After this pack is committed and staging migrations run cleanly, you're ready for **Phase 1: Data Migration** — exporting your Google Sheets and Dance Infusion JSON into the new schema. We'll handle that in the next Claude Code session.

## Rollback if something goes wrong

If a migration fails partway through or you need to start over on staging:

```sql
-- DANGER: drops all tables, views, types in the public schema
drop schema public cascade;
create schema public;
grant usage on schema public to anon, authenticated, service_role;
grant all on schema public to postgres;
```

Then re-run migrations 001 through 012 from scratch. Do not run this on production once it has real data.
