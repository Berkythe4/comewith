---
name: project_email_conversations
description: Email-any-actor/venue feature + Conversations threads (logged, team-visible, restrictable); bounce tracking
metadata:
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

Email + Conversations system shipped 2026-06-25 (migration 056, commit 4abe2b1). Email any actor/venue from where it makes sense; every send becomes a logged thread.

**Entry points (dashboard.html):** shared `openCompose(recipients, ctx)` modal. Wired from: Actors screen (per-row ✉ `data-aemail` + multi-select checkboxes `data-asel`/`data-asel-all` + "✉ Email selected"), Vendors screen (`data-ven-vsel` multi + `data-ven-email` + vendor profile), Venues list + venue profile (`data-ven-email="venue:<id>"`), event hub People tab (`data-hub-psel` multi + `data-hub-email-person`/`data-hub-email-people`). Compose: subject gets `[<source>]` prefix, body gets a deep link `?goto=<kind>&id=<id>` (handled on load → `gotoSource()`), visibility = team | restricted (+ pick who).

**Backend:** `conversations` / `conversation_messages` / `conversation_acl` tables (056) with RLS via `can_see_conversation()` = master OR creator OR (visibility='team' AND `user_can_access_module('conversations')`) OR in ACL. New **Conversations** module in module_registry (Audience group, signed_off, roles operations/marketing/full). Edge fn **`send-actor-email`** (admin-only JWT): resolves actor/venue email, creates a thread + outbound message per recipient, sends via Resend (FROM `Come With <berky@comewith.org>`, **reply_to berky@comewith.org**), stores `resend_id`. Supports replying into a thread (`conversation_id`). `resend-webhook` extended: correlates delivery/bounce events by `resend_id` → updates message status AND logs a visible "⚠ Delivery failed — bounced" event message in the thread.

**Conversations screen** (`loadConversations`/`convState`): thread list (status badge, 🔒 restricted, ⚠️ bounce flags, filters) → thread view (`openThread`) with message bubbles + status, **Send reply** (emails), **Add internal note** (`note` direction, no email), visibility control + `convManageAcl` (manage who can see), "Go to source".

**Verified end-to-end on prod** (then cleared): send to keith.berkman@gmail.com (delivered) + bounced@resend.dev (→ bounced, logged in thread); actor_id email resolution; **team vs restricted visibility across martin + henry logins** (henry sees team threads, NOT restricted, and is 403-blocked from posting to a restricted thread); reply; note. Bounce-back works because the Resend webhook IS configured + the comewith.org domain IS verified (arbitrary-recipient sends succeed).

**RESIDUAL / known gap:** capturing inbound human REPLIES automatically needs inbound email (Resend inbound domain + MX records — external DNS step, not done). For now replies go to **berky@comewith.org's inbox** (reply_to); paste them into the thread via "Add internal note" to keep the back-and-forth logged. Wiring an inbound webhook to auto-append is the future step. See [[project_email_campaigns]] (same Resend stack), [[project_staff_access_model]] (the module gate this rides on).
