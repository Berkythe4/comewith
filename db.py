"""Run SQL against the Supabase project via the Management API.

Reads credentials from environment variables (never written to disk):
  SBP_PAT  - Supabase Personal Access Token (sbp_...)
  SBP_REF  - Project ref (the subdomain piece of the dashboard URL)

Usage:
  python db.py "select 1 as one;"                       (inline SQL)
  python db.py supabase/migrations/013_grants.sql       (SQL file)
  python db.py -                                        (SQL from stdin)

Prints the API response as indented JSON to stdout.
Prints a one-line "project= source=" banner to stderr so you can confirm
which Supabase project the query hit.

Exits nonzero on HTTP/network error and prints the response body to stderr.
Uses only the Python stdlib so no pip install is required.
"""

import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path


def load_dotenv():
    """Read .env in the script directory and populate os.environ for any
    keys not already set. Tolerates a missing file. Lines starting with #
    and blank lines are skipped. Values may be wrapped in single or double
    quotes; quotes are stripped."""
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_dotenv()

def normalize_ref(raw):
    """Accept the bare ref, the project URL, or the dashboard URL and return
    just the short ref the Management API expects."""
    if not raw:
        return raw
    s = raw.strip().rstrip("/")
    s = s.replace("https://", "").replace("http://", "")
    if ".supabase.co" in s:
        s = s.split(".supabase.co")[0]
    elif "supabase.com/dashboard/project/" in s:
        s = s.split("supabase.com/dashboard/project/")[1].split("/")[0]
    return s


PAT = os.environ.get("SBP_PAT")
REF = normalize_ref(os.environ.get("SBP_REF"))

if not PAT or not REF:
    print(
        "ERROR: Set SBP_PAT and SBP_REF environment variables before running.\n"
        "  PowerShell:  $env:SBP_PAT='sbp_xxx'; $env:SBP_REF='your-project-ref'\n"
        "  cmd:         set SBP_PAT=sbp_xxx && set SBP_REF=your-project-ref",
        file=sys.stderr,
    )
    sys.exit(1)

if len(sys.argv) < 2:
    print(
        "Usage: python db.py <SQL string | path/to/file.sql | ->",
        file=sys.stderr,
    )
    sys.exit(1)

arg = sys.argv[1]

if arg == "-":
    sql = sys.stdin.read()
    source = "<stdin>"
else:
    path = Path(arg)
    if path.exists() and path.is_file():
        sql = path.read_text(encoding="utf-8")
        source = str(path)
    else:
        sql = arg
        source = "<inline>"

URL = f"https://api.supabase.com/v1/projects/{REF}/database/query"
HEADERS = {
    "Authorization": f"Bearer {PAT}",
    "Content-Type": "application/json",
    # Cloudflare in front of the Management API rejects urllib's default UA
    "User-Agent": "comewith-db.py/1.0 (+https://comewith.org)",
}

print(f"[db.py] project={REF} source={source}", file=sys.stderr)

body = json.dumps({"query": sql}).encode("utf-8")
req = urllib.request.Request(URL, data=body, headers=HEADERS, method="POST")

try:
    with urllib.request.urlopen(req, timeout=180) as resp:
        raw = resp.read().decode("utf-8")
        result = json.loads(raw) if raw else None
except urllib.error.HTTPError as e:
    err_body = e.read().decode("utf-8", errors="replace")
    print(f"FAIL on {source}", file=sys.stderr)
    print(f"HTTP {e.code} {e.reason}", file=sys.stderr)
    print(f"Body: {err_body}", file=sys.stderr)
    sys.exit(1)
except urllib.error.URLError as e:
    print(f"FAIL on {source}: network error: {e.reason}", file=sys.stderr)
    sys.exit(1)

print(json.dumps(result, indent=2, default=str))
