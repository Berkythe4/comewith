---
name: reference-machine-toolchain
description: "This machine cannot run node --check or the anon exposure sweep, and its .env points db.py at prod by default"
metadata: 
  node_type: memory
  type: reference
  originSessionId: 9285a14c-2927-4aa5-9b36-c63f3a5610ad
  modified: 2026-08-22T01:29:30.540Z
---

Verified 2026-08-21 on this machine (`C:\Users\keith\comewith`):

- **No JS runtime at all** — `node`, `deno`, `bun`, `npx` are all absent. The
  `node --check` loop CLAUDE.md documents for `dashboard.html` cannot run here.
  Workaround used: extract the inline module and parse it with Python `esprima`,
  shimming ES2019-2022 syntax (`||=`, `?.[`, `catch {`, top-level `await`), with
  the **pre-edit file as a control** and a deliberately broken copy as a negative
  control. Installing Node would remove the whole workaround.
- **`.env` has no `SUPABASE_PROD_PUBLISHABLE_KEY`**, so
  `scripts/check_anon_exposure.py` exits with "no prod URL / publishable key".
  Grants can still be verified in SQL with `has_table_privilege()` /
  `has_function_privilege()`, but that does **not** exercise PostgREST end to
  end. Henry's machine has the same gap.
- **`.env` contains a bare `SBP_REF=yaytdosxfhcqatmhctzk` — prod.** CLAUDE.md
  says not to have one, so that the target project is visible in the command
  being approved. Until it is removed, a bare `python db.py file.sql` silently
  targets production. Always pass the literal
  `SBP_REF=yaytdosxfhcqatmhctzk python db.py ...` anyway.

Verify each of these still holds before relying on it — they are machine
configuration and may have been fixed.
