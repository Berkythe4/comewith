"""End-to-end test of Social Calendar v2 + the social-mcp connector, on PROD.

    python scripts/e2e_social_mcp.py

Drives the DEPLOYED connector over real MCP (the same HTTP claude.ai sends) and
replays the dashboard's own writes as Keith - `set local role authenticated` with
his JWT claims - so RLS and 217's claude_caption trigger are both in the path.

  1. create_post_skeleton            -> idea, owned by Janelle
  2. update_post_draft (caption)     -> drafted, claude_drafted_at set; list shows it
                                        and it reads as needing review (the 🤖 rule)
  3. "Use this" as Keith              -> Final caption filled from Claude's
  4. Approve as Keith                 -> stage approved + note; the bell insert passes
                                        RLS (run in a ROLLED-BACK txn, so Janelle is
                                        not pinged by a test)
  5. posted + results as Keith        -> get_results returns them
Plus the refusals, live: bad secret, the final caption, an approved post, a
delete, and Keith writing claude_caption (the trigger).

The test post is DELETED at the end (notes cascade). connector_log rows stay -
it is a log. Needs SBP_PAT and SOCIAL_MCP_SECRET in .env.
"""
import json
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REF = "yaytdosxfhcqatmhctzk"
KEITH = "bd71f88d-fd0a-4e3a-a712-7ac958318c8b"
JANELLE = "52245d29-ca9f-4ad7-8dcd-6c893459dd3b"
STAMP = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
TITLE = f"E2E TEST {STAMP} - delete me"
fails = 0


def env(name):
    for raw in (ROOT / ".env").read_text(encoding="utf-8").splitlines():
        if raw.strip().startswith(name + "="):
            return raw.split("=", 1)[1].strip().strip("'\"")
    return None


SECRET = env("SOCIAL_MCP_SECRET")
URL = f"https://{REF}.supabase.co/functions/v1/social-mcp"


def check(cond, label):
    global fails
    print(("PASS  " if cond else "FAIL  ") + label)
    if not cond:
        fails += 1


def mcp(method, params=None, url=None):
    body = json.dumps({"jsonrpc": "2.0", "id": 1, "method": method, "params": params or {}}).encode()
    req = urllib.request.Request(url or f"{URL}/{SECRET}", data=body, method="POST", headers={
        "content-type": "application/json", "accept": "application/json, text/event-stream",
        "mcp-protocol-version": "2025-06-18"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            text = r.read().decode("utf-8")
            status = r.status
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")
    line = next((l for l in text.splitlines() if l.startswith("data:")), None)
    return status, json.loads(line[5:] if line else text)


def tool(name, args):
    _, res = mcp("tools/call", {"name": name, "arguments": args})
    if "error" in res:
        return {"_error": res["error"]}
    r = res["result"]
    text = r["content"][0]["text"] if r.get("content") else ""
    try:
        data = json.loads(text)
    except ValueError:
        data = {"_text": text}
    data["_isError"] = bool(r.get("isError"))
    return data


def sql(q):
    p = subprocess.run([sys.executable, str(ROOT / "db.py"), "-"], input=q, capture_output=True,
                       text=True, encoding="utf-8", env={**__import__("os").environ, "SBP_REF": REF}, cwd=ROOT)
    if p.returncode != 0:
        return {"_error": (p.stderr or p.stdout).strip()[-600:]}
    out = p.stdout.strip()
    return json.loads(out) if out else []


def as_keith(body):
    """Run statements as Keith, signed in, then COMMIT."""
    return sql(f"""begin;
select set_config('request.jwt.claims', '{{"sub":"{KEITH}","role":"authenticated"}}', true);
set local role authenticated;
{body}
commit;""")


def main():
    if not SECRET:
        sys.exit("SOCIAL_MCP_SECRET is not in .env")
    pid = None
    try:
        # ---- refusals that need no post ------------------------------------
        st, _ = mcp("tools/list", url=f"{URL}/not-the-secret-not-the-secret-not-the-x")
        check(st == 404, "bad secret -> 404")
        st, res = mcp("tools/list")
        names = sorted(t["name"] for t in res["result"]["tools"])
        check(names == ["add_note", "create_post_skeleton", "get_results", "list_posts", "log_results", "update_post_draft"],
              "exactly six tools, none of them a delete")

        # ---- 1. skeleton ---------------------------------------------------
        r = tool("create_post_skeleton", {"title": TITLE, "scheduled_at": "2026-10-20T18:00:00-04:00",
                                           "account": "come_with", "format": "reel", "phase": "radio",
                                           "brief": "E2E brief"})
        pid = r.get("created", {}).get("id")
        check(bool(pid) and r["created"]["stage"] == "idea" and r.get("owner") == "janelle@comewith.org",
              "1. create_post_skeleton -> stage idea, owner Janelle")
        dup = tool("create_post_skeleton", {"title": TITLE.lower(), "scheduled_at": "2026-10-20T09:00:00-04:00",
                                             "account": "come_with", "format": "reel", "phase": "radio"})
        check(dup["_isError"] and "already exists" in json.dumps(dup), "   duplicate title on the same day refused")
        if not pid:
            return

        # ---- 2. Claude drafts ----------------------------------------------
        r = tool("update_post_draft", {"id": pid, "claude_caption": "Claude's E2E caption ✨ #ComeWith"})
        check(not r["_isError"] and r["updated"]["stage"] == "drafted" and r["moved_to_drafted"],
              "2. update_post_draft -> claude_caption written, idea -> drafted")
        row = sql(f"select stage, claude_caption, claude_drafted_at is not null as stamped, caption, owner_id::text "
                  f"from social_posts where id = '{pid}'")[0]
        check(row["stamped"] and row["caption"] is None and row["owner_id"] == JANELLE,
              "   claude_drafted_at set; final caption untouched (null)")
        lst = tool("list_posts", {"from": "2026-10-20", "to": "2026-10-20"})
        mine = [p for p in lst.get("posts", []) if p["id"] == pid]
        check(len(mine) == 1 and mine[0]["claude_caption"] and not mine[0]["final_caption"],
              "   list_posts shows it with claude_caption and no final -> the list's 🤖 rule is true")
        r = tool("update_post_draft", {"id": pid, "caption": "connector writing the final"})
        check(r.get("_isError") or "_error" in r, "   refused: connector touching the final caption")
        r = as_keith(f"update social_posts set claude_caption = 'Keith typed this' where id = '{pid}';")
        check("_error" in r and "only by the Claude connector" in r["_error"],
              "   refused: a signed-in user writing claude_caption (217 trigger)")

        # ---- 3. Use this (the editor's patch) --------------------------------
        r = as_keith(f"update social_posts set caption = claude_caption where id = '{pid}';")
        row = sql(f"select caption = claude_caption as same from social_posts where id = '{pid}'")[0]
        check("_error" not in r and row["same"], "3. Use this -> Final caption filled from Claude's, as Keith")

        # ---- 4. Approve ----------------------------------------------------
        r = as_keith(f"""update social_posts set stage = 'approved', copy_status = 'approved' where id = '{pid}';
insert into social_post_notes (post_id, body) values ('{pid}', '— Approved —');""")
        row = sql(f"select stage, (select count(*) from social_post_notes n where n.post_id = p.id) as notes "
                  f"from social_posts p where id = '{pid}'")[0]
        check("_error" not in r and row["stage"] == "approved" and row["notes"] == 1,
              "4. Approve -> stage approved, note in the thread")
        bell = sql(f"""begin;
select set_config('request.jwt.claims', '{{"sub":"{KEITH}","role":"authenticated"}}', true);
set local role authenticated;
insert into notifications (user_id, from_user_id, kind, title, body, subject_type, subject_id)
values ('{JANELLE}', '{KEITH}', 'approved', 'Approved: {TITLE}', 'Approved — ready to schedule.', 'post', '{pid}');
reset role;
select count(*)::int as bells from notifications where subject_id = '{pid}' and user_id = '{JANELLE}' and kind = 'approved';
rollback;""")
        check(isinstance(bell, list) and bell and bell[0]["bells"] == 1,
              "   the bell to the owner passes RLS as Keith (rolled back: Janelle is not pinged by a test)")
        r = tool("update_post_draft", {"id": pid, "claude_caption": "too late"})
        check(r["_isError"] and "only edits posts at idea or drafted" in json.dumps(r),
              "   refused: connector editing an approved post")
        r = tool("delete_post", {"id": pid})
        alive = sql(f"select count(*)::int n from social_posts where id = '{pid}' and deleted_at is null")[0]["n"]
        check((r.get("_isError") or "_error" in r) and alive == 1, "   refused: delete attempt (no such tool; post still there)")

        r = tool("log_results", {"id": pid, "views": 1})
        check(r["_isError"] and "only be logged on a posted post" in json.dumps(r),
              "   refused: log_results on a post that is not posted (approved)")

        # ---- 5. posted + results -------------------------------------------
        r = as_keith(f"""update social_posts set stage = 'posted', posted_at = now(),
  results = '{{"views": 1234, "likes": 56, "shares": 7, "saves": 8}}'::jsonb where id = '{pid}';""")
        today = datetime.now().strftime("%Y-%m-%d")
        res = tool("get_results", {"from": today, "to": today})
        got = [p for p in res.get("posts", []) if p["id"] == pid]
        check("_error" not in r and len(got) == 1 and got[0]["results"] == {"views": 1234, "likes": 56, "shares": 7, "saves": 8},
              "5. posted + results entered -> get_results returns them")

        before = sql(f"select md5(concat_ws('|', title, caption, stage, scheduled_for::text, owner_id::text, "
                     f"claude_caption, brief, account, format, phase)) h from social_posts where id = '{pid}'")[0]["h"]
        r = tool("log_results", {"id": pid, "views": 2000, "comments": 9})
        row = sql(f"select results, md5(concat_ws('|', title, caption, stage, scheduled_for::text, owner_id::text, "
                  f"claude_caption, brief, account, format, phase)) h from social_posts where id = '{pid}'")[0]
        check(not r["_isError"] and row["results"] == {"views": 2000, "likes": 56, "comments": 9, "shares": 7, "saves": 8},
              "log_results on a posted post -> merged into results (views updated, comments added, the rest kept)")
        check(row["h"] == before, "   log_results wrote nothing but results")
        r = tool("log_results", {"id": pid, "views": 1, "caption": "sneak"})
        check(r.get("_isError") or "_error" in r, "   refused: log_results with another field")

        # ---- notes + log ---------------------------------------------------
        r = tool("add_note", {"id": pid, "text": "E2E: Claude was here."})
        who = sql(f"select author_name, author_id is null as no_profile from social_post_notes "
                  f"where post_id = '{pid}' and body = 'E2E: Claude was here.'")
        check(not r["_isError"] and who and who[0]["author_name"] == "Claude" and who[0]["no_profile"],
              "add_note -> in the thread, signed Claude")
        log = sql(f"select tool, ok from connector_log where post_id = '{pid}' order by id")
        tools_ok = {(x["tool"], x["ok"]) for x in log}
        check(("create_post_skeleton", True) in tools_ok and ("update_post_draft", False) in tools_ok
              and ("add_note", True) in tools_ok and ("delete_post", False) in tools_ok
              and ("log_results", True) in tools_ok
              and sum(1 for x in log if x["tool"] == "log_results" and not x["ok"]) >= 2
              and sum(1 for x in log if x["tool"] == "update_post_draft" and not x["ok"]) >= 2,
              f"connector_log has every call, schema rejections and delete attempts included ({len(log)} rows)")
    finally:
        if pid:
            sql(f"delete from social_posts where id = '{pid}';")
            left = sql(f"select count(*)::int n from social_posts where id = '{pid}'")
            print(f"\ncleanup: test post deleted ({'ok' if left and left[0]['n'] == 0 else 'CHECK MANUALLY ' + pid})")
    print("\nALL PASS" if not fails else f"\n{fails} FAILED")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
