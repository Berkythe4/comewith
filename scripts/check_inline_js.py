"""Syntax-check the inline <script> of an HTML page, with a control.

CLAUDE.md's rule for `dashboard.html` is "syntax-check by extraction, not by
re-reading — and run the same extraction against the pre-edit version as a
control, otherwise an artifact of the extraction itself reads as a real error
introduced by the edit." This is that, as a command.

There is no JS runtime on every machine that works this repo (the laptop has
neither node nor deno), so the parser is `esprima` — pure Python, already
installed, and ES2017-era. Newer syntax is therefore DOWNGRADED before parsing:
optional chaining, nullish coalescing, logical assignment, optional catch
binding. The rewrite changes SEMANTICS, never syntax, and it is applied to both
sides — so anything it cannot handle fails identically in the control and in the
edited file, which is the signal that it is an artifact and not your bug.

`dashboard.html` currently ends on a top-level `await`, which esprima also
predates. So the expected clean result for that file is BOTH sides failing on
the same final construct — the parse having reached the end of the module is the
pass. dj.html and the smaller pages parse clean outright.

Usage:
  python scripts/check_inline_js.py dj.html
  python scripts/check_inline_js.py dashboard.html

Exit code is 0 when the edited file parses as far as the control did, 1 when the
edit made it worse.

⚠ Do not write patch scripts for these files through a shell heredoc. Backslash
escapes get processed on the way in even inside a quoted heredoc, so a `\\n` meant
for a JS string arrives as a real newline and silently breaks a string literal —
which is exactly the failure this script was extended to catch (2026-09-12).
Use the editor tools for anything containing escapes.
"""

import io
import re
import subprocess
import sys

try:
    import esprima
except ImportError:  # pragma: no cover
    sys.exit("esprima is not installed:  python -m pip install esprima")


def js_of(text):
    """The page's inline module body. These pages carry exactly one."""
    m = re.search(r"<script(?:\s[^>]*)?>(.*)</script>", text, re.S)
    if not m:
        return None
    return m.group(1)


def downgrade(js):
    # Logical assignment first, so ??= is gone before ?? is rewritten.
    for op in ("??=", "||=", "&&="):
        js = js.replace(op, "=")
    js = re.sub(r"\?\.(?=[\[\(])", "", js)          # a?.[x] / a?.(x) -> a[x] / a(x)
    js = js.replace("?.", ".")
    js = js.replace("??", "||")
    js = re.sub(r"catch\s*\{", "catch (_e) {", js)  # optional catch binding
    return js


def check(label, text):
    """Returns (ok, line) — line is where it stopped, 0 when it parsed clean."""
    body = js_of(text)
    if body is None:
        print(f"{label}: no inline <script> found")
        return False, -1
    js = downgrade(body)
    try:
        esprima.parseModule(js)
        print(f"{label}: parsed clean ({len(js)} chars)")
        return True, 0
    except Exception as e:
        m = re.search(r"Line (\d+)", str(e))
        line = int(m.group(1)) if m else -1
        print(f"{label}: stopped at line {line} — {e}")
        for i, text_line in enumerate(js.split("\n")):
            if line - 3 <= i + 1 <= line + 1:
                print(f"   {i + 1:>6} {text_line[:160]}")
        return False, line


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    path = sys.argv[1]
    edited = io.open(path, encoding="utf-8").read()
    head = subprocess.run(["git", "show", f"HEAD:{path}"], capture_output=True)
    control = head.stdout.decode("utf-8", "replace")

    ok_ctl, line_ctl = (None, None)
    if control.strip():
        ok_ctl, line_ctl = check("control (HEAD)", control)
    else:
        print("control (HEAD): not in git — nothing to compare against")

    ok_new, line_new = check("edited      ", edited)

    if ok_new:
        return 0
    if ok_ctl is False:
        # Both stopped. The edit is only implicated if it stopped EARLIER — the
        # line numbers move when an edit adds lines, so "same or later" is the
        # pass, and the shared construct is the artifact.
        if line_new >= line_ctl:
            print("\nOK — both sides stop on the same construct (parser artifact, not this edit).")
            return 0
        print("\nFAIL — the edited file stops EARLIER than the control. That is this edit.")
        return 1
    print("\nFAIL — the control parsed clean and the edited file does not.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
