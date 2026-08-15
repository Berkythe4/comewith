"""Deploy a Supabase Edge Function to prod via the Management API.

The CLI is linked to STAGING, and as of CLI 2.101.0 it also rejects the newer
`sbp_v0_...` personal-access-token format outright ("Invalid access token
format"), so `supabase functions deploy` is not usable here at all. The
Management API accepts the same token fine.

Reads SBP_PAT from .env (never printed). Preserves the function's existing
verify_jwt setting unless --no-verify-jwt is passed.

Usage:
  python scripts/deploy_edge_function.py pull-dice pull-ticketmaster
  python scripts/deploy_edge_function.py --ref <ref> <slug> [<slug>...]

Exits nonzero on the first failure.
"""

import json
import mimetypes
import sys
import urllib.error
import urllib.request
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PROD_REF = "yaytdosxfhcqatmhctzk"
API = "https://api.supabase.com"
UA = "comewith-deploy/1.0 (+https://comewith.org)"


def load_env():
    env = {}
    p = ROOT / ".env"
    if not p.exists():
        return env
    for raw in p.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def multipart(fields, files):
    """fields: {name: str}; files: [(name, filename, content_type, bytes)]."""
    boundary = "----comewith" + uuid.uuid4().hex
    out = bytearray()
    for name, value in fields.items():
        out += f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n'.encode()
        out += value.encode("utf-8") + b"\r\n"
    for name, filename, ctype, data in files:
        out += (
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"; '
            f'filename="{filename}"\r\nContent-Type: {ctype}\r\n\r\n'
        ).encode()
        out += data + b"\r\n"
    out += f"--{boundary}--\r\n".encode()
    return bytes(out), f"multipart/form-data; boundary={boundary}"


def call(method, url, pat, data=None, ctype=None):
    headers = {"Authorization": f"Bearer {pat}", "User-Agent": UA}
    if ctype:
        headers["Content-Type"] = ctype
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=180) as resp:
        raw = resp.read().decode("utf-8")
        return json.loads(raw) if raw else None


def deploy(slug, ref, pat):
    src = ROOT / "supabase" / "functions" / slug / "index.ts"
    if not src.exists():
        print(f"FAIL {slug}: no {src}", file=sys.stderr)
        return False

    # Keep whatever the live function already declares — a deploy should ship new
    # code, not silently change who is allowed to call it.
    try:
        cur = call("GET", f"{API}/v1/projects/{ref}/functions/{slug}", pat) or {}
    except urllib.error.HTTPError:
        cur = {}
    verify_jwt = cur.get("verify_jwt", True)
    if "--no-verify-jwt" in sys.argv:
        verify_jwt = False

    meta = {
        "name": cur.get("name") or slug,
        "entrypoint_path": "index.ts",
        "verify_jwt": verify_jwt,
    }
    body, ctype = multipart(
        {"metadata": json.dumps(meta)},
        [("file", "index.ts", mimetypes.types_map.get(".ts", "text/typescript"),
          src.read_bytes())],
    )
    url = f"{API}/v1/projects/{ref}/functions/deploy?slug={slug}"
    try:
        res = call("POST", url, pat, data=body, ctype=ctype) or {}
    except urllib.error.HTTPError as e:
        print(f"FAIL {slug}: HTTP {e.code} {e.read().decode('utf-8', 'replace')[:400]}", file=sys.stderr)
        return False
    print(f"OK   {slug} -> version {res.get('version')} status {res.get('status')} verify_jwt={verify_jwt}")
    return True


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    ref = PROD_REF
    if "--ref" in sys.argv:
        ref = sys.argv[sys.argv.index("--ref") + 1]
        args = [a for a in args if a != ref]
    if not args:
        print(__doc__, file=sys.stderr)
        sys.exit(1)

    pat = load_env().get("SBP_PAT")
    if not pat:
        print("ERROR: SBP_PAT not in .env", file=sys.stderr)
        sys.exit(1)

    print(f"[deploy] project={ref}", file=sys.stderr)
    ok = True
    for slug in args:
        ok = deploy(slug, ref, pat) and ok
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
