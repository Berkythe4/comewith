# Claude memory — repo snapshot

Claude Code's memory lives **outside the repo**, in a per-machine, per-path folder
(`~/.claude/projects/<path-slug>/memory/`). It does **not** sync between machines.
A second machine — Keith's laptop — therefore starts with none of it, and would
rebuild project state from git archaeology instead of fact.

This folder is a **snapshot** of the desktop's memory, committed so any machine
can read it. `MEMORY.md` is the index: one line per memory, newest concerns
first. Read the index, then open only the files you need.

## Rules

- **The live memory on the desktop is the original; this is a copy.** When they
  disagree, the live one is newer.
- **Re-snapshot at every session close** (it's step 6 of the close routine in
  `SESSION_CLOSE_PROMPTS.md`), so the gap is never more than one session:
  ```
  cp ~/.claude/projects/C--Users-Admin-Documents-Comewith/memory/*.md DEV_DOCS/claude-memory/
  ```
  On the laptop the path slug differs — it's derived from the checkout path.
- **Memories record what was true when written.** If one names a file, function,
  column or flag, verify it still exists before acting on it.
- **Never put a credential in a memory file** — this folder is in git now.
  Snapshot date: 2026-08-15 (60 files, scanned clean).

## Working from the laptop

There's no way to import these into the laptop's own Claude memory, and no need
to: `CLAUDE.md` points here, so they're read as ordinary docs. If the laptop
becomes the main machine, snapshot *its* memory here at close and keep the same
one-way discipline — do not merge two divergent memory sets by hand.
