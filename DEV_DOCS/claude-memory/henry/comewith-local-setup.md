---
name: comewith-local-setup
description: "State of the local C:\\comewith dev environment — what's set up and what's still missing."
metadata: 
  node_type: memory
  type: project
  originSessionId: 43d45cf6-b868-4f32-9a6c-033e1bc425b5
  modified: 2026-08-22T00:36:31.446Z
---

As of 2026-08-15, `C:\comewith` is a fresh clone of `github.com/Berkythe4/comewith`
(branch `master`, not `main`).

**Toolchain is complete** as of 2026-08-15, all resolving **by bare name**: `python`
3.12.10, `gh` 2.97.0, `node` v24.19.0, `git` 2.55.0. `gh` is authed as `hjzaradich`,
wired as git's credential helper via `gh auth setup-git`, and `viewerPermission` on the
repo is **WRITE**.

**A tool that "isn't on PATH" here is usually a stale process env, not shadowing.**
Claude Code captures PATH at launch, so anything installed mid-session resolves against
the OLD PATH. The WindowsApps `python.exe` stub was never shadowing the real install —
the user PATH already ordered `Python312\` ahead of `WindowsApps\`. Diagnose by composing
a fresh PATH from the Machine + User env vars and resolving against that; if that finds
the right binary, the fix is a **restart**, not clearing App Execution Aliases. Dotfiles
don't help — the Bash tool does not source `~/.bashrc` — and `${PATH}` expansion in
settings.json `env` is undocumented, so setting PATH there risks replacing it outright
and losing `git`.

**Prod access works** as of 2026-08-15. `.env` holds `SBP_PAT` + `SBP_REF_PROD`; Henry's
Supabase account was invited to the **Come With** org, so the PAT now sees both
`comewith-prod` (`yaytdosxfhcqatmhctzk`) and `comewith-staging` (`qjosjafbizxdtkisyrqm`).
Management API calls need a **browser User-Agent** or Cloudflare answers 403.

**`db.py` takes arbitrary SQL — there is no read/write split.** Inline SQL and migration
files go through one code path, and the Management API query endpoint accepts
multi-statement SQL, so `select …; drop …` would run both. That is why **no
`Bash(python db.py:*)` pattern can be safely allowlisted**, not even a `select`-prefixed
one. It imports stdlib only, so no pip install is needed. It reads `SBP_REF`, which
`.env` deliberately does NOT define — pass it explicitly so the target project is visible
in the command being approved.

**Pass the ref LITERALLY: `SBP_REF=yaytdosxfhcqatmhctzk python db.py "…"`.** CLAUDE.md
writes the convention as `SBP_REF=$SBP_REF_PROD`, but that **does not work from the Bash
tool** — `SBP_REF_PROD` lives only in `.env`, which `db.py` reads for itself and which is
never sourced into the shell. So `$SBP_REF_PROD` expands to the empty string, `SBP_REF=""`
lands in `os.environ`, `load_dotenv()` skips the key because it is already present, and
db.py exits with the misleading "Set SBP_PAT and SBP_REF" (it names both even when only
one is missing). The literal form is also what the allowlist prefix matches, so it is the
only form that runs without a prompt.

**`db.py` against PROD is allowlisted as of 2026-08-15.** `.claude/settings.local.json`
carries `Bash(SBP_REF=yaytdosxfhcqatmhctzk python db.py:*)`, added by Henry who wants the
workflow automated. The prod ref is baked into the prefix on purpose: any other project
still prompts. It is a standing grant over arbitrary SQL — `db.py` has no read/write
split — so the care now lives in the migration, not the approval dialog. **Dry-run every
migration first** by copying it with `commit;` swapped for `rollback;` and running that:
it caught a nested-window-function error in 142 that no amount of re-reading had.
Before this rule, every call was blocked by the classifier — batching the checks into
single UNION ALL statements (`supabase/checks/pre_apply.sql` / `post_apply.sql`) dates
from then and is still worth doing, since one call is one round trip.

**The classifier will not let Claude widen its own permissions** — both the direct
settings edit and the `update-config` skill are blocked. A permission rule has to be
added by hand.

The publishable anon key is public and lives in the frontend — `radio.html:282`
(`CWURL` / `CWKEY`) — so the anon-401 invariant check needs no PAT at all.

**`scripts/check_anon_exposure.py` CANNOT run on this machine** (confirmed 2026-08-21):
it wants `SUPABASE_PROD_PUBLISHABLE_KEY` in `.env`, which here holds only `SBP_PAT` and
`SBP_REF_PROD`, so it exits `FAIL no prod URL / publishable key in .env`. Until that key
is added, do the anon checks by hand with `curl`, reading the key out of the frontend
(`grep -ohE "sb_publishable_[A-Za-z0-9_-]+" dashboard.html | head -1`). **Prove the key
works before trusting any 401** — that script exists precisely because an empty key
answers 401 for everything and made a real leak look blocked. Use
`v_public_events?select=*` and check for a **200 with a non-empty body**; don't pick a
column at random, that view has no `id` and a bad column name returns 400, which reads
like a broken key.

**Still missing:** **ffmpeg/ffprobe** — blocks verifying video renders by pulling a frame.

`dashboard.html` is ~1.3 MB in one inline `<script type="module">`. Never read it into
context; the convention is now written up in `CLAUDE.md`. Syntax-check by extracting that
block to a `.mjs` and running `node --check`, **and run the same extraction against the
pre-edit version as a control** so an extraction artifact can't be misread as a real error.

**The Browser pane cannot open `file://` URLs**, so the dashboard can't be loaded locally
for a console check.

**Why it matters:** `master` auto-deploys to comewith.org through Netlify, so pushing to
`master` is a production deploy, not a merge. Work on a branch and let Keith merge.

**The `C:\comewith` checkout MOVES under you mid-session.** Other sessions and machines
share it: over one session on 2026-08-15 it went `docs/dashboard-editing-convention` →
`fix/post-apply-checks` → `feature/calendar-focus-scope` → `master`, and PRs were merged
while work was still in flight. Never assume your branch is still checked out — re-check
`git branch --show-current` before editing, and if it has moved, do NOT switch it back
(that yanks the tree out from under whoever is using it). Use a throwaway worktree
instead: `git worktree add <scratchpad>/wt-x <branch>`, edit there, commit, push,
`git worktree remove`. This worked cleanly four times in a row and never disturbed the
other session. Corollary: a branch you pushed may be merged before you push a follow-up,
so **check PR state before assuming you can add a commit to it** — a merged PR needs a
fresh branch cut from the current `master`, not another push to the old one.
