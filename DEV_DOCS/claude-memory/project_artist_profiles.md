---
name: project_artist_profiles
description: Public artist profiles + collective + content tagging + artist self-service editing, plus homepage ticker/DI-button/community changes
metadata:
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

Built 2026-06-26 (commits 9f682c6 → 4abca61). Two parts:

**Homepage content changes:** removed the "daytime community" section. Ticker is
now content-driven (`strip.items`, default `Music · Rave · Community · Daytime ·
Dance Infusion · Brooklyn` — no House/Disco). Dance Infusion impact card has
separate **Tickets** + **Donate** buttons, each with a show/hide checkbox + label
+ URL, all in the **Site Editor** (keys `di.ticket_show/_label/_url`,
`di.donate_*`; falls back to one "Get a ticket or donate" link when no URLs set).
Site Editor gained a "Ticker strip" section + checkbox rendering for `*_show`
keys (SITE_BOOLS).

**Artist profiles (migration 065):** `actors` gained `bio, photo_path, soundcloud,
tiktok, public_profile, collective_rank, edit_token`. Public anon views:
`v_public_artists` (collective = public_profile actors), `v_artist_gigs` (from
event_participants, public/completed events only), `v_artist_content` (unnests
events.recap_videos where each item has an `artist_id`). Seeded public_profile=true
for existing dj/artist actors.
- **artist.html?id=<actor_id>** — public profile: photo, bio, socials (ig/sc/tiktok/
  web), Content grid (tagged recap media w/ lightbox), gig history.
- **Homepage collective** (`#djs`/loadCollective) loads v_public_artists → clickable
  avatar chips → artist.html.
- **Dashboard Artists tab**: click an artist → profile editor (public toggle,
  collective order, bio, socials, photo upload/remove). "On site" column + KPI.
- **Content tagging**: each recap-video row in the event editor has an artist
  dropdown (RECAP_ARTISTS); stored as `artist_id` in the recap_videos item. Existing
  videos auto-tagged by label match (7 items across 5 artists).
- **Self-service**: `artist-self` edge function (token=actors.edit_token, deployed
  --no-verify-jwt; actions get/save/photo/photo_remove via service role) +
  **artist-edit.html?token=** (no-login page: bio/socials/photo, HEIC handled).
  Dashboard "Copy update link" / "Email update link" (via send-notice).

**Artist INTAKE / create (added 2026-07-07, commit 8d9d73c):** parallel to the update
flow. `artist-intake` edge fn (public, --no-verify-jwt) creates an `actors` row + `artist`
role from a submission, auto-generates edit_token, keeps `public_profile=false` for admin
review, dedupes by email (re-submit returns existing edit link), optional base64 photo,
honeypot (`company`/`hp` field), pings berky@ via Resend on new intake. **artist-intake.html**
= public onboarding form (name/socials/bio/photo), returns the artist their private
artist-edit.html link on success. **Dashboard Artists tab "＋ New artist" button** → modal:
create directly (sb insert actors+actor_roles; RLS = `can_see_people()`, lands in the profile
editor to add photo + toggle live) OR copy/email the public intake link (send-notice).
`openNewArtist`/`createArtist`/`sendIntakeLink` in dashboard.html. E2E verified on prod
(create/dedupe/honeypot/validation + returned edit link resolves via artist-self; test rows cleaned).

Original build E2E verified on prod: views anon-readable, function get/save/404, content+gigs
return per artist, all pages 200. Open: socials empty until filled; DI button URLs
empty until set in Site Editor. See [[project_customer_frontend]], [[project_user_management]].
