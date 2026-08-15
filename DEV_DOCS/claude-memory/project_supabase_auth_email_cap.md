---
name: project-supabase-auth-email-cap
description: Prod has NO custom SMTP, so auth emails are capped at 2 per hour project-wide — breaks magic links and silently throttles listener signups
metadata:
  type: project
---

Prod auth config (checked 2026-07-30): **`smtp_host = None`** — still on Supabase's
built-in shared sender, whose limit is **`rate_limit_email_sent = 2` per hour,
PROJECT-WIDE** (not per user). `smtp_max_frequency = 60`.

**What it broke:** Martin couldn't sign in — a magic link to `martin@comewith.org`, a
second to his krneyentertainment address ("rate exceeded"), then a password reset that
was never sent. His account was fine all along: confirmed, `sub_admin`, not
deactivated, `has_password: true`, and previously signed in on 2026-06-26. **The
immediate unblock is email+password at the dashboard login, not a magic link.**

**The bigger problem:** `radio.html:369` creates listener accounts with
`signInWithOtp` — the free "one tap, no password" flow. Those are the same auth emails
under the same 2/hour cap, so **any listener signing up after the second one in an hour
silently gets nothing** and the site cannot tell them why. Same for the password reset
at `dashboard.html:2246`.

**Fix (needs Keith, ~2 min in the Supabase dashboard):** Project Settings → Auth →
SMTP: `smtp.resend.com:587`, user `resend`, password = the Resend API key, sender
`no-reply@comewith.org`. comewith.org is already verified in Resend for campaigns.
Then raise `rate_limit_email_sent` to 30–100/hour — Supabase only allows that once
custom SMTP exists. Do it in the dashboard UI rather than pasting the API key into a
chat.

Do NOT send a magic link to an address with no auth user (e.g. the krney one):
signups aren't disabled, so it creates a junk `customer`-role account.
