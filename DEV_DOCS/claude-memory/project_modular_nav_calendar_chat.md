---
name: modular-nav-calendar-chat
description: "096/097 (2026-07-17) — modular product nav (7 sellable groups), master Calendar & Tasks module, team chat w/ realtime; all applied to prod + committed a4f2920"
metadata: 
  node_type: memory
  type: project
  originSessionId: 18764a6f-2571-469b-8c44-dd3618fac89e
---

Applied to prod + committed (a4f2920, 2026-07-17), per Keith's "modular sellable tool" direction:

- **Nav (096)**: `module_registry.nav_group` regrouped into product modules — Pinned (calendar) / Workflow / Finance / Marketing / Venues / Artists / Radio / Team HQ. Sidebar renders collapsible groups (state in `localStorage.cw_nav_collapsed`, active group auto-opens); `NAV_GROUP_ORDER` in dashboard.html must match nav_group values. `actors` relabeled "People & Orgs".
- **Calendar & Tasks** (`calendar` module, pinned): month grid + agenda; layers = our events (colored by series), ⭐ milestones (`tasks.milestone`, added in 096 — the ONLY tasks on the grid, per Keith's clutter concern), social posts, RA/TM market shows, roster-artists-elsewhere (name match vs `ra_events.lineup`). Multi-select filters series/status/venue/artist. Tasks board below grid = cross-event task list w/ filters/sort/quick-add.
- **Team chat (097)**: `chat_channels`/`chat_members`/`chat_messages`, kinds team/dm/event; DMs deduped by `dm_key` = sorted "uid:uid"; member-only DM privacy (master has NO implicit read of others' DMs). Realtime publication on `chat_messages` (first realtime use in the app) + 45s poll fallback. 💬 header button on every tab; Users tab has ✉/💬 per user; team emails mirror into the DM as kind='email' rows via `openCompose` ctx.onSent hook.
- **RLS lesson (bit us here)**: `INSERT..RETURNING` enforces the SELECT policy mid-statement, where a security-definer helper that re-queries the table cannot see the new row → "new row violates RLS". Fix: put `created_by = auth.uid()` directly in the SELECT policy, not only inside the helper. Also: RLS can be smoke-tested on prod via Management API with `set_config('request.jwt.claims', ...)` + `set local role authenticated` inside BEGIN..ROLLBACK (grant temp tables to authenticated).
- Related: [[project-event-hub-and-di2-money]], [[project-user-management]], [[project-email-conversations]].

Pushed to origin/master 2026-07-17/18 (a4f2920 … 8101e57 close-out) → Netlify auto-deploy. Close-out done 2026-07-18: workflow map (+📅 Calendar, +💬 Team comms steps), APIs map (Realtime note), ROADMAP "reconciled 2026-07-18" section, CLAUDE.md RLS RETURNING rule + deactivation contract, build stamp updated. Residual: inbound email replies still land in berky@ (unchanged).

**Later same-day follow-ups (…af5acd1)**: shared openEditTask modal (calendar board + hub tasks + milestone chips); hub "✉ Email task list" (buildTasksEmailHtml, inline-styled email w/ overdue/status sections, message field, live preview) — sends with **send-actor-email `single_thread: true`** (ONE conversation, all recipients on one To line; fn deployed via `supabase functions deploy --project-ref` + SBP_PAT as access token) and posts a kind='email' notice into # Team chat w/ "Open thread →" + "✓ Complete" (delete row = done; conversation permanent). Conversations threads: "📥 Log received email" — paste or .eml upload parsed client-side (no mailbox access). [hidden] gotcha: any CSS display rule beats the hidden attribute — chat panel/dock/badges need `[hidden]{display:none!important}`.

**098 follow-ups (same day, bcdfde5)**: chat minimize → docked pill w/ unread count (`cw_chat_min`); user deactivation = `profiles.deleted_at` + is_admin/is_master_admin/user_can_access_module all guard `deleted_at is null` (instant full revoke, verified on prod); Users tab checkboxes + "Email selected" multi-send (one conversation per person + DM mirror); Site Editor/Review moved to Team HQ. Deactivated users stay in get_team_members (no filter) so master can reactivate; they're dropped from loadStaff pickers.
