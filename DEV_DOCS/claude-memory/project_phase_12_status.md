---
name: project-phase-12-status
description: Phase 12 (2026-05-29) — placeholder homepage swap + admin email/password auth shipped to prod
metadata: 
  node_type: memory
  type: project
  originSessionId: d97501fd-d79e-460c-80ae-ea0889c23091
---

Phase 12 closed 2026-05-29 on prod (comewith.org), two changes, both pushed to origin/master:

**(1) Homepage swapped to placeholder.** Root `index.html` is now a simple landing page (logo, "Brooklyn, NY", Events panel = "announced soon", Contact panel = email/Instagram/bookings). Decision: a placeholder beats a half-built site. The prior half-built public form (full Supabase inquiry + newsletter wiring) was archived to `legacy/index-v2-publicform.html` — restore from there when finishing it. Placeholder has a deliberately subtle top-right "Sign in" link → `dashboard.html`; low contrast and the ~900KB embedded base64 logo are intentional/acceptable for a placeholder (user declined contrast bump + logo extraction). Commit `9ab25ee`.

**(2) Admin email+password auth shipped** in `dashboard.html`. Admins (master_admin + sub_admin) can now sign in with email+password OR magic link interchangeably. Magic link kept as the fallback (byte-for-byte unchanged) — blank password field sends a link. Always-on "Set password" link in the sidebar footer does first-time-set AND change via `updateUser({ password })` (no `has_password` flag — would drift). Customer auth ([[project-phase-6-status]] index/portal, sign.html) stays magic-link/token only — NOT touched. Commit `6f233b8`.

Netlify auto-deploys root statics from master (no netlify.toml/_redirects; `index.html` = live homepage). Also fixed `origin` remote URL casing → `https://github.com/Berkythe4/comewith.git`. Builds on [[project-phase-11-status]].
