---
name: project-phase-11-status
description: Phase 11 (production cutover) closed 2026-05-28. comewith.org is live on Supabase prod. Migration done. Multiple known issues + redesign items deferred to future sessions.
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 11 closed 2026-05-28. The full Supabase migration (Phases 0-11)
shipped in a single day. comewith.org is now served by Netlify from
the new v2 codebase, talking to `comewith-prod` Supabase project.

## What's live on prod

### Pages (canonical names after the v2 rename)
- `index.html` — public inquiry form + mailing list + hero/about/services
  /Series we run/Watch sections
- `dashboard.html` — admin dashboard (read + write for inquiries,
  agreements, clients, income, expenses, equipment, events, sponsors,
  sponsorships, artists, subscribers, campaigns)
- `customer_portal.html` — customer login + agreements list +
  re-open signing link; auto-redirects admins to dashboard.html
- `sign.html` — customer agreement signing page
- `confirm.html` — mailing list double-opt-in confirmation
- `unsubscribe.html` — mailing list unsubscribe
- `legacy/` — all Apps Script-era pages preserved here

### Edge Functions on prod (yaytdosxfhcqatmhctzk)
send-agreement · get-agreement-by-token · mark-signed ·
inquiry-notify · get-event-hub · subscribe ·
confirm-subscription · unsubscribe · send-campaign ·
resend-webhook

### Prod secrets set
- RESEND_API_KEY ✓
- SITE_URL = https://comewith.org ✓

### Prod Auth config
- Site URL: https://comewith.org
- Redirect allow-list: https://comewith.org/** + https://www.comewith.org/**

### Migrations applied to prod
- 001 through 014 (35 tables, RLS on all, 6 storage buckets,
  2 cron jobs)

## Known issues to fix in a future session

### 🔴 Berky's role on prod
Berky signed in as `berky@comewith.org` on prod via the customer
portal, but profiles.role defaulted to 'customer' (the prod DB is
fresh; the Phase 2 role escalation only happened on staging). He
landed on the customer portal instead of the admin dashboard.

**Fix:** one SQL update on prod (via `SBP_REF=yaytdosxfhcqatmhctzk
python db.py "..."`):
```sql
update public.profiles
set role = 'master_admin'
where email = 'berky@comewith.org';
```

### 🟡 Apps Script triggers possibly still enabled
User did NOT confirm they disabled Apps Script triggers in
script.google.com. Form submissions to old endpoints (if any
URLs exist in the wild) may still be processed by the legacy
Apps Script → Sheets pipeline in parallel with the new Supabase
flow.

**Fix:** script.google.com → Come With project → Triggers (clock
icon) → delete each trigger.

### 🟡 events/dance-infusion-2/index-v2.html still points at staging
Per the DI hub publish gate, this file was NOT renamed/swapped
during the cutover. When the publish gate is lifted, it needs:
- Hardcoded staging URL+key swapped to prod values (same swap
  done in the 6 root files in sprint 4)
- Optionally renamed events/dance-infusion-2/index-v2.html →
  events/dance-infusion-2/index.html if you want a clean URL

### 🟡 Resend webhook not configured
RESEND_WEBHOOK_SECRET is not set on prod and Resend dashboard
doesn't have the webhook URL configured. resend-webhook will
accept unsigned events (because no secret is set) but Resend
isn't sending them yet.

**Fix:** Resend dashboard → Webhooks → Add endpoint:
- URL: https://yaytdosxfhcqatmhctzk.supabase.co/functions/v1/resend-webhook
- Events: email.delivered, email.bounced, email.complained
- Copy signing secret → `supabase secrets set RESEND_WEBHOOK_SECRET=whsec_xxx --project-ref yaytdosxfhcqatmhctzk`

### 🟡 Scheduled campaign sends (Phase 10 deferral)
pg_cron can't easily invoke send-campaign because it requires
an admin JWT. See project_phase_10_status memory for the two
candidate solutions (cron-secret header vs vault + service_role).

## Redesign items deferred (user flagged at close)
User stated "we need to redesign a lot of things" — explicit
items mentioned during this session:
- Landing page layout / messaging needs more iteration
- Sign-in routing (admin landed on customer portal because of
  the role issue above — but the routing logic is correct, just
  needs the role to be set)
- Possibly other UX tweaks discovered as Berky uses the live site

These aren't blocking — comewith.org is functional. Treat as a
backlog the user will pull from as they use the system.

## Hardcoded TODO callouts now resolved
All Phase 5-10 "TODO(phase-11)" markers in Edge Function code
are gone — replaced with SITE_URL env var lookup. The Phase 11
sprint 2 commit refactored these.

## Time tracking
Estimated ~25-30 min wall-clock. Came in around ~25 min for
sprints 1-4, plus ~10 min of Phase 11.5 landing-page iteration,
plus user hands-on time for verification.

## What this session accomplished
Came in this morning with:
- Apps Script + Google Sheets backed comewith.org
- 12 SQL migrations sitting in a folder, unapplied
- "Phase 2 is next" status

Ends with:
- comewith.org running on Supabase + Netlify + 10 Edge Functions
  + Resend
- Berky has a real admin dashboard with write flows for the
  daily ledger work he used to do in Sheets
- Public inquiry, mailing list, customer signing flow, all live
- Dance Infusion #2 event data seeded into the data lake
- pg_cron jobs running nightly maintenance
- All Apps Script-era code preserved in /legacy/ for reference

Related: [[project-phase-10-status]],
[[project-di-hub-publish-gate]], [[feedback-time-estimates]]
