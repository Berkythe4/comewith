# Henry's machine — Claude memory snapshot

Separate folder, deliberately. The files one level up are the **desktop's** 60-odd
memories and `../MEMORY.md` is the desktop's index. Henry's machine has its own, much
smaller memory (Claude Code memory is per-machine and does not sync), and the merge
routine's `cp ~/.claude/projects/<slug>/memory/*.md DEV_DOCS/claude-memory/` would have
**overwritten the desktop's 64-entry index with a 3-entry one**. Hence the namespace.

Snapshotted 2026-08-15 (Strategy rebuild close). Scanned for credentials before commit:
these mention variable NAMES (`SBP_PAT`, `SBP_REF_PROD`) and the prod project ref, which
is already public throughout the repo — no token values.

If a third machine starts shipping, give it a folder here too rather than merging indexes.
