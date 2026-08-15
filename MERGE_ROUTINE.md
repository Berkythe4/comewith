# The Merge Routine — Come With

*(Called the "session close" until 2026-08-15, and renamed because that's not what it
is any more. Three machines now ship into this repo — Keith's desktop, Keith's laptop,
Henry's — and every close is also a merge. Old file name: `SESSION_CLOSE_PROMPTS.md`.)*

Paste one of the prompts below at the end of any session where you shipped real work.
Pick the variant by session size. Adapted from the AI Planner's protocol; the
Come-With-specific bits are the prod safety checks (anon-revoked financial views,
migration-vs-prod drift) instead of the Planner's test suite / ADR pipeline.

---

## Step 0 — merge FIRST, before you write a word of docs

Other machines have been shipping. Do this before anything else, every time:

```
git fetch origin && git log --oneline HEAD..origin/master
```

- **Pull before you document.** CARRYOVER, LEARNINGS and ROADMAP are all
  append-heavy shared files; writing them against a stale base is how you get a
  conflict in the one file whose job is to tell the next person what's true.
- **Check the migration numbers.** Take the next free number *after* pulling —
  2026-08-15 had `140_site_owner.sql` written as `138` while the laptop was
  simultaneously landing `138_notes_assignment.sql` and `139_notes_to_tasks.sql`.
  Two migrations sharing a number is a merge conflict in prod, not in git.
- **Re-verify prod facts you gathered before the pull.** Another machine may have
  applied a migration that moves them.
- **Preserve every "This session shipped" block.** They stack, newest first. Nobody's
  session gets overwritten because someone else closed later the same day.

---

## When to use this

- **Always** at the end of a session that shipped a migration or touched prod data.
- **Always** when a new design decision was made (it belongs in `LEARNINGS.md`).
- **Always** when `ROADMAP.md` parked/backlog items changed.
- Optional for a tiny copy-only or single-line fix.

---

## Every close, regardless of size (added 2026-08-15)

Two machines now work this repo (desktop + laptop), and Claude Code's memory is
per-machine — it does **not** sync. So a close isn't finished until the *repo*
carries the state:

- **Re-snapshot Claude memory** into `DEV_DOCS/claude-memory/`:
  `cp ~/.claude/projects/<path-slug>/memory/*.md DEV_DOCS/claude-memory/`
  (the slug is derived from the checkout path, so it differs per machine).
  Scan for credentials before committing — that folder is public in the repo.
- **Say in CARRYOVER which machine the session ran on**, and name any branch the
  work is parked on, with the exact command to pick it up.
- **`master` auto-deploys to Netlify.** Never merge un-green-lit work to master
  to "tidy up" a close. Docs-only commits on master are fine — they rebuild the
  site with identical content.

---

## Quick close (small session — one fix / one UI tweak, no schema change)

Use when: a bugfix or single UI change shipped; no migration; no new decision.

```
You are running a Quick Close for the Come With session that just shipped.

What shipped this session (fill in):
- [one-line summary]

Do these steps exactly:
1. Verify prod invariant (read-only, anon REST via .env): all 5 financial views return 401 — v_event_summary, v_kpi_event_financials, v_kpi_parties, v_kpi_dance_infusion, v_kpi_dashboard. If any returns 200, STOP and flag (blanket anon-grant regression — see LEARNINGS §4 / migration 019).
2. Report git state: confirm the work is committed and pushed (or note what's held).
3. Append a one-line entry to CARRYOVER.md "This session shipped" with what changed. Do NOT rewrite the rest of CARRYOVER.
4. Report briefly: what you appended + the anon-401 result.

Then STOP. No new files, no LEARNINGS edits, no ROADMAP changes unless explicitly requested.
```

---

## Standard close (medium session — a feature, 1–2 migrations, some decisions)

Use when: a feature shipped with one or more migrations; a decision or two was made; CARRYOVER + ROADMAP need updating.

```
You are running a Standard Close for the Come With session that just shipped.

What shipped this session (fill in):
- [feature / scope summary]
- [migrations added — list numbers]
- [decisions made — short list]

Do these steps exactly:
1. VERIFY PROD INVARIANTS (read-only, via Supabase Management API + anon REST using .env keys; never print secrets):
   - All 5 financial views return anon 401 (list in Quick close step 1). If any 200 → STOP and flag.
   - supabase/migrations/ matches prod: report the highest migration number and confirm it's applied (no drift).
2. UPDATE CARRYOVER.md — refresh "State summary" (migration range, latest LEARNINGS §, roles, git) and "Tomorrow's default"; update "This session shipped" and "Parked / next".
3. APPEND TO LEARNINGS.md — for any NEW decision, append a new numbered §N at the END with date + rationale. Preserve every existing section; never edit in place (supersede if needed).
4. UPDATE ROADMAP.md — tick off what shipped; add/clear backlog + parked items.
5. CREATE reviews/session_YYYY-MM-DD.md — SHORT (a few bullets, not an essay): what happened, the ARC (what you set out to do vs. what you discovered), key decisions, parked/next. This preserves the narrative CARRYOVER loses when it's overwritten.
6. Report briefly: what you updated, migration range, anon-401 result.

Then STOP. Push only when Keith says so.
```

---

## Full close (big session — multi-part build, 3+ migrations, multiple decisions)

Use when: a major build shipped across several commits/migrations; multiple decisions; ROADMAP needs a real refresh.

```
You are running a Full Close for the Come With session(s) that just shipped.

What shipped (fill in):
- [build name + the parts/commits]
- [migrations added — list numbers]
- [key decisions — 3–7 bullets]
- [anything reconciled on prod (data fixes, etc.)]

Do these steps exactly:

PART 0 — Pre-close checks (read-only)
- All 5 financial views return anon 401. If any 200 → STOP and flag (LEARNINGS §4 / 019).
- supabase/migrations/ matches prod (no drift); report highest applied number.
- Confirm the live site deployed (poll comewith.org/dashboard.html for the shipped markup) and report git state (committed/pushed/held).

PART 1 — Durable docs
- LEARNINGS.md: append a new §N per NEW decision (date + rationale). Append-only; preserve + supersede, never delete.
- CARRYOVER.md: refresh in full — State summary (prod, migration range, roles, financial-view status, latest LEARNINGS §, git), Tomorrow's default, This session shipped, Parked / next.
- ROADMAP.md: move shipped items to a "Done" note; refresh Parked (design-first) + Backlog.
- reviews/session_YYYY-MM-DD.md: a SHORT narrative (a few bullets, not an essay) — what happened, the ARC (set-out-to-do vs. discovered), key decisions, parked/next, one honest note. This is the durable record of the session's story; CARRYOVER is a snapshot that gets overwritten.

PART 2 — Conventions check
- Re-read CLAUDE.md; if the session established a new standing rule (RLS, grants, series contract, scope), add it to CLAUDE.md as well as LEARNINGS.md.

PART 3 — Report
- Short summary: what shipped, what's parked, what's next, and any open risk (e.g. placeholder data, deferred decisions).

Then STOP. Push only when Keith says so.
```

---

## How to use

1. End the build session.
2. In a fresh session (or after `/clear`), paste the appropriate template.
3. Fill in the "What shipped" placeholder with the real deliverables.
4. Run. Review the diffs. Commit/push on confirm.

Without this, real session discoveries and the live prod state drift out of the docs.

---

## Trigger triggers (when to interrupt yourself with a close)

- **"I'll just ship one more migration"** → close first if 2+ migrations since the last close.
- **"I'll document it tomorrow"** → the decision rationale fades; capture it in LEARNINGS now.
- **A new standing rule came up in conversation** → ALWAYS close: it needs to land in both LEARNINGS.md and CLAUDE.md, or the next session won't honor it.
- **Touched prod data / grants** → ALWAYS close with the anon-401 verification, so a regression can't ride along silently.

---

## What the close protocol buys

The 016/017 anon-grant regression (financial views silently re-exposed) was caught only
because post-apply verification re-checked the 401 invariant. Baking that check into every
close means a grant regression can never survive a session. The rest — CARRYOVER, LEARNINGS,
ROADMAP staying true to prod — is what lets the next session start from fact instead of
reconstructing state from git archaeology.
