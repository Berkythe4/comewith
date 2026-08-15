---
name: project_site_review_audit
description: Site Review module (076) — in-app audit log under Site Editor; 2026-07-08 full-site audit results + open decisions + prioritized API plan
metadata: 
  node_type: memory
  type: project
  originSessionId: 6ae7a3fb-5f8f-4348-8f22-cc711985cf02
---

Full-site audit run 2026-07-08 (5 parallel review agents + prod DB checks), commit `b27f25c`,
migration **076 applied to prod**.

**Site Review module** (key `site-review`, Insights group, sort 194 — right under Site Editor;
`signed_off=false` → master-only): renders `site_review_items` as grouped tables
(bug/improvement/capability/data/saved) with inline status select
(open/review/planned/fixed/dismissed) + add-item modal. Code: `loadSiteReview`/`srSetStatus`/
`srOpenAdd` in dashboard.html. 19 findings seeded. This is the standing maintenance backlog —
future audits should ADD ROWS here, not create new docs.

**Security verified clean**: all 7 financial views 401 for anon; zero RLS-enabled-no-policy
tables; zero RLS-disabled tables; no orphaned FKs; series contract intact (4 valid values only).

**Fixed in the sweep**: convertInquiry (email dedupe + checked role/status writes); equipment
load-in checkbox reverts on failed save; send-agreement surfaces failed status-flip as `warning`
(redeployed); og:image/og:url/twitter:card added to index/watch/artist (og:image pinned to the
CURRENT logo storage URL — re-uploading the logo requires updating 3 page heads); portal empty
state links berky@; fee-to-expense error copy.

**Site Editor gained a "Dashboard settings" (ops.*) section** — first key `ops.ra_guestlist_type`
(RA guest-list export Type column). Pattern for future in-app settings: seed a `site_content`
`ops.*` key + `SITE_LABELS` entry; read it where needed.

**2026-07-08 PM — Keith triaged, 8/8 planned items EXECUTED (077, commit 030f788):** rate
limiting (subscribe 10-min confirm throttle via subscribers.confirm_sent_at; inquiry-notify
3/hr cap), public-endpoint error sanitization (admin fns keep detail BY DESIGN), FROM_EMAIL/
REPLY_TO_EMAIL secrets (set on prod, read by all 9 senders), **email_templates table + editor on
the Templates screen** (keys: artist_update_link, artist_intake_invite, subscribe_confirm,
survey_invite; senders fill {{placeholders}}, hardcoded fallback; agreement email deliberately
NOT templated — structured body), vendor categories via ops.vendor_categories, social series ×4,
homepage fallback DJ names REMOVED (consent risk), 7 stale pre-Supabase HTML tools pruned to
untracked archive/old-reports (LIVE impact-report.html/public-audit.html + AUDIT_TABLES kept —
those live under events/dance-infusion/di-02-2026-05/reports/). 12 fns redeployed; throttle +
sanitized errors verified live. **Miss Vee merge done** (Victoriarose soft-deleted; roles
artist/contractor/dj/donor; backup backups/premerge_2026-07-08_missvee.json). Venues: Knicks
G5→Crossroads, Henry Showcase→1163 Putnam. Review log: 17 fixed / 1 open (og:image pin —
awaiting Keith's understanding/decision) / 1 review (KPI by-design). July 4th Weekend venue
still unset. NOTE: actors.status only allows 'active'/'on_hold' (no 'archived').

**Known-deliberate (do NOT "fix")**: KPI views cover only Parties+DI (production/content use the
Events-page money models); homepage PAST/DJS arrays are offline fallbacks; social-post series
dropdown only offers Parties/DI; archive impact-report copies under events/di-02.../ are stale
on purpose (live report is Supabase-driven).

**API plan (in the 🔌 APIs map, priority order)**: #1 Instagram Graph (pull-instagram-stats
built — Keith needs Meta app + IG_USER_ID/IG_ACCESS_TOKEN secrets), #2 GA4 (property + service
acct → new pull-ga4-stats fn → metric_snapshots), #3 self-hosted ICS calendar feed from
v_public_events (no keys, ~2 hrs), #4 TikTok (when channel active), #5 Eventbrite/DICE (only if
ticketing moves), #6 Spotify (embeds sooner, API later). No-API with CSV workarounds built:
RA (import + guest-list export), Partiful, Simplifi.

Workflow map gained Artist intake + Site review steps. See [[project_artist_profiles]],
[[project_resend_engagement_tracking]], [[project_api_integrations]].
