---
name: reference-machine-toolchain
description: "This machine cannot run node --check; the anon sweep now works here, and .env still points db.py at prod by default"
metadata:
  node_type: memory
  type: reference
  originSessionId: 9285a14c-2927-4aa5-9b36-c63f3a5610ad
  modified: 2026-08-27T20:51:40.531Z
---

Machine configuration for `C:\Users\keith\comewith` (the laptop). Verify each
before relying on it — this is config, and config gets fixed.

- **No JS runtime at all** — `node`, `deno`, `bun`, `npx` all absent (still true
  2026-08-27). The `node --check` loop CLAUDE.md documents for `dashboard.html`
  cannot run here. Workaround: extract the inline module, downlevel the syntax
  esprima cannot see (`||=`, `??`, `?.`, `catch {`, numeric separators), and
  **wrap the body in an async IIFE after lifting the `import` lines out** —
  esprima has no top-level `await`, which is the one that blocked this before.
  Run it three ways: git HEAD as control (must PASS, or the checker is the
  problem), the working copy, and a copy with a brace deleted (must FAIL, or the
  checker would pass a real bug). Script kept at `scratchpad/syntax_check.py`;
  worth moving into `scripts/` if this machine stays Node-less.
- **`SUPABASE_PROD_PUBLISHABLE_KEY` IS now in `.env`** (added 2026-08-22), so
  `scripts/check_anon_exposure.py` and `scripts/check_financial_views.py` both
  run here. The key was never secret — it ships inside `dashboard.html` because
  the browser needs it. **Both scripts read `.env` directly and ignore the
  process environment**, so setting the variable on the command line does
  nothing; it has to be in the file. Henry's machine can be fixed the same way.
- **`.env` still contains a bare `SBP_REF=yaytdosxfhcqatmhctzk` — prod.**
  CLAUDE.md says not to have one, so the target project is visible in the command
  being approved. Until it is removed, a bare `python db.py file.sql` silently
  targets production. Pass the literal `SBP_REF=yaytdosxfhcqatmhctzk python db.py …`
  anyway — it is also the form Henry's allowlist prefix matches.
- **`git fetch` failed once with `libcurl-4.dll` blocked by an Application
  Control policy**, then succeeded on retry; `git ls-remote` worked throughout.
  If a fetch dies that way, retry before concluding the remote is unreachable.

Related: [[project-fpa-planning-tool]]
