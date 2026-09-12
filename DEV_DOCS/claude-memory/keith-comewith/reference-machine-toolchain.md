---
name: reference-machine-toolchain
description: "No JS runtime here (use scripts/check_inline_js.py), heredocs corrupt scripts three ways, the anon sweep works, and .env still points db.py at prod by default"
metadata:
  node_type: memory
  type: reference
  originSessionId: 9285a14c-2927-4aa5-9b36-c63f3a5610ad
  modified: 2026-09-12T17:21:01.444Z
---

Machine configuration for `C:\Users\keith\comewith` (the laptop). Verify each
before relying on it — this is config, and config gets fixed.

- **No JS runtime at all** — `node`, `deno`, `bun`, `npx` all absent (still true
  2026-09-12). The `node --check` loop CLAUDE.md documents for `dashboard.html`
  cannot run here. **This is now solved and committed: `python
  scripts/check_inline_js.py <file>`** (added 2026-09-12). It extracts the inline
  module, downlevels what esprima cannot parse (`||=` `&&=` `??=`, `?.`, `??`,
  `catch {`), and runs the same extraction against `git show HEAD:<file>` as a
  control. It does NOT solve top-level `await` — `dashboard.html` ends on one, so
  the expected pass for that file is **both sides stopping on the same final
  line**, and the exit code knows that. `dj.html` parses clean outright.
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
- **In the esprima downlevel, rewrite `?.[` and `?.(` BEFORE `?.`.** Doing `?.`
  first turns `a?.[k]` into `a.[k]`, which is a syntax error esprima reports as
  "Unexpected token [" — and it looks exactly like a real error introduced by the
  edit. Order: `?.[` → `[`, `?.(` → `(`, then `?.` → `.`, then `??` → `||`.
  Also verified 2026-09-02: **check the extraction boundaries return a non-empty
  block** (`src.index(A):src.index(B)`), because if A appears after B the slice is
  empty and esprima happily reports PARSE OK on nothing.
- **A `<<'EOF'` heredoc in the Bash tool is NOT literal — an apostrophe in the
  body breaks the whole command** with ``unexpected EOF while looking for matching
  `'``, pointing at a line number inside the heredoc. The command looks like it is
  wrapped in outer single quotes, so quoted-heredoc semantics do not protect the
  body. It bites on ordinary prose (`Keith's`, `the page's`) and on SQL comments,
  and it bites *after* nothing has been written. Balanced quotes (`', '` in SQL)
  are fine, which is why some heredocs work and hide the rule. **Write prose and
  comment-heavy files with the Write tool**, and keep heredocs for
  apostrophe-free content. Verified 2026-08-31 on two separate failures.
- **`git fetch` failed once with `libcurl-4.dll` blocked by an Application
  Control policy**, then succeeded on retry; `git ls-remote` worked throughout.
  If a fetch dies that way, retry before concluding the remote is unreachable.

Related: [[project-fpa-planning-tool]]

- **A long SQL string cannot be passed to `db.py` as an argument here.** Windows
  caps a command line near 32k, so a full-station rewrite (~100 statements) dies
  with `FileNotFoundError: [WinError 206] The filename or extension is too long`
  before the process starts. The failure is safe (nothing ran) but looks like a
  missing file. `db.py` accepts a FILE PATH as its argument -- write the SQL to a
  temp .sql and pass that. `Radio/render/apply_station_from_rekordbox.py` does
  this above 8k.
- **Bash heredocs mangle non-ASCII on this machine.** A `<<'PYEOF'` block
  containing an em-dash silently failed a string match against a UTF-8 file on
  2026-09-10. Write the script with the Write tool and run the file instead.
- **Bash heredocs also eat BACKSLASH ESCAPES, quoted or not — the worst of the
  three, because it succeeds.** On 2026-09-12 a `<<'PY'` patch script wrote `\\n\\n`
  into a JS `confirm()` string; Python received `\n\n` already unescaped and put a
  REAL newline inside a single-quoted JS literal, breaking `dashboard.html`. The
  same run turned a regex `r"\\bcatch"` into a literal backspace character, so the
  substitution silently never matched. Both were found by eye, not by any check.
  **Any patch script containing a backslash goes through the Write tool, never a
  heredoc** — this now covers apostrophes, non-ASCII AND escapes, so the honest
  rule is simply: do not write scripts through heredocs on this machine.
