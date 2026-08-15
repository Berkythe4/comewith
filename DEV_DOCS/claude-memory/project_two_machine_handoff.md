---
name: project-two-machine-handoff
description: "Keith works this repo from a desktop AND a laptop; Claude memory doesn't sync, so state lives in the repo — CARRYOVER.md + DEV_DOCS/claude-memory/ snapshot, re-snapshotted at every close"
metadata: 
  node_type: memory
  type: project
  originSessionId: 6264aafb-c0a6-4e5c-bdca-a6f2c2a39988
  modified: 2026-08-15T20:34:52.417Z
---

Set up 2026-08-15. Keith added a **laptop** alongside the desktop. Claude Code's
memory is per-machine and per-path (`~/.claude/projects/<slug>/memory/`) and does
**not** sync — so a laptop session starts blind unless the state is in the repo.

The channel is the repo, in this order:
- **`CARRYOVER.md`** — where the last session left off + "Parked / next". Now names
  which machine the session ran on, and any branch work is parked on.
- **`DEV_DOCS/claude-memory/`** — a committed snapshot of the desktop's 60 memory
  files (index `MEMORY.md`, plus a README on the one-way rule). Scanned clean of
  credentials before committing; keep it that way.
- **`CLAUDE.md`** — gained a "Start of session" / "End of session" section, so it's
  auto-loaded on either machine and points at all of the above.
- **`SESSION_CLOSE_PROMPTS.md`** — the close ritual, plus a new "every close" block:
  re-snapshot memory, name the machine, never merge un-green-lit work to master.

**Why:** the desktop's memory was the only record of months of decisions, and none
of it was reachable from a fresh checkout.

**How to apply:** at every session close, re-copy the memory folder into
`DEV_DOCS/claude-memory/` before committing. Live memory on the machine you're on is
newer than the snapshot; never hand-merge two divergent snapshots. And remember
`master` auto-deploys to Netlify — held work goes on a branch named in CARRYOVER.
Related: [[project-radio-discovery-window]].
