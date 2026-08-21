---
name: project_social_calendar
description: Social content calendar module — migration 044 applied to prod 2026-06-23; Calendar/List/Board + threaded notes (timeline REMOVED 2026-08-21); invite-user fn authored not deployed
metadata: 
  node_type: memory
  type: project
  originSessionId: 8c7b42a9-8c6e-4ec8-8ee5-50c34f2040a4
---

Social content calendar for Keith <-> Janelle collaboration. Built 2026-06-23,
merged to master + deployed to comewith.org via Netlify (merge commit 8fa2a62).

**Migration 044 (APPLIED to prod 2026-06-23, additive/non-financial):**
- `social_posts` (title, caption, channels[], series, content_pillar, stage,
  scheduled_for, posted_at, owner_id, event_id, link_url, asset_url, asset_status,
  cta, soft-delete). stage ∈ idea|drafted|review|planned|scheduled|posted|archived.
- `social_post_notes` (post_id cascade, author_id default auth.uid(), body,
  created_at) — the timestamped conversation thread.
- RLS on both = `public.user_can_access_module('social-calendar')`; notes
  edit/delete restricted to author or master. New leaf tables, no Events-hub
  coupling (so real RLS was safe to apply here, unlike 042).
- Flipped module_registry `social-calendar` to built=true, signed_off=true so
  marketing (Janelle) sees it. See [[project_staff_access_model]].

**dashboard.html:** `loadSocialCalendar()` renders 3 views. **CORRECTED 2026-08-21:**
the Timeline view (a chronological vertical spine, `timelineCardHtml`/`.tl-*`) was
**deleted** and the List view rebuilt to the events-list spec — `data-table cc-table`,
`ccPostBand()` banding, `ccTitle()` names with rename behind the pencil, and stage /
scheduled date / channels / pillar editable in place via `socialPatch()`. The two
single-value filter dropdowns were replaced by multi-select chips in `#socialFilters`
(`social.fStage` / `fSeries` / `fChan` / `q`, all arrays). Cycle is Calendar → List →
Board. `channels` is an array and `content_pillar` is free text, so neither takes a
plain select — see LEARNINGS §38. Post modal (reuses `openKpi`) does full CRUD + a threaded notes
panel (`loadPostNotes`/`addPostNote`) that timestamps each note. Team tab gained
an "＋ Add person" button. Stage colors: `SOCIAL_STAGE_COLOR`.

**Share snapshot (export):** "📤 Share snapshot" toolbar button → `exportSnapshot()`
builds a self-contained read-only HTML doc (current filters, posts + full notes
threads, branded) and opens it in a new tab to Print/Save-as-PDF or save — so
Keith can send a point-in-time snapshot to Janelle without her logging in.
Popup-blocked → falls back to file download. No live link / no auth involved.

**invite-user Edge Function (AUTHORED, NOT deployed):** master-gated; verifies
caller JWT is master_admin, then service-role `inviteUserByEmail` + sets
role/staff_role. Deploy: `SUPABASE_ACCESS_TOKEN=$SBP_PAT supabase functions
deploy invite-user --project-ref yaytdosxfhcqatmhctzk`. CAVEAT: until 043 is
applied, a new sub_admin can read financial views via direct REST — apply 043
before inviting non-finance staff.

To get Janelle actually using it: (1) merge branch to master (Netlify deploy),
(2) deploy invite-user, (3) invite Janelle as marketing. Each is a deliberate
user step.
