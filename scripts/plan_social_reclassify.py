"""Plan the 2026-10-02 social-post re-classification from a backup, rules only.

    python scripts/plan_social_reclassify.py backups/social_posts_pre_reclassify_2026-10-02.json

Prints every change (before -> after), every flag, and the past DI posts left
alone, and writes the plan as JSON next to the backup. Migration 219 is built
from that plan by id, so what the dry run shows is exactly what executes.

Rules (Keith's sprint spec, 2026-10-02). LIVE POSTS ONLY - deleted rows are
dropped before any rule runs. Never touches caption, stage, scheduled_for, owner.
  1 Radio (all dates): title has CWR or radio (case-insensitive) OR series =
    'Come With Radio' -> phase radio; account come_with only if come_with/null;
    format only if null: 'recap' -> carousel; 'release' or Ep<n> -> reel; else flag.
  2 Upcoming DI (NY date >= today): whole-word DI / DI3 / Dance Infusion ->
    account collab; phase awareness < 2026-11-14, event on it, post after.
  3 Specific posts override 2 (by id, or exact title + NY date).
"""
import json
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

DI_DAY = "2026-11-14"
RADIO_RX = re.compile(r"cwr|radio", re.I)
DI_RX = re.compile(r"\b(di3?|dance infusion)\b", re.I)
EP_RX = re.compile(r"\bep\s?\d+", re.I)
DI_N_RX = re.compile(r"\bdi\d+\b", re.I)

SPECIFIC = [
    {"id": "c3a36e7b-0896-42ee-a36c-f0fe7d702086", "expect_title": "Dance Infusion Official Annoncement",
     "set": {"title": "DI3 Official Announcement", "format": "carousel", "phase": "awareness", "account": "collab"}},
    {"id": "75ec2bbc-bb97-4da9-96c5-e8a03723001b", "expect_title": "DI b2b post (Kristen/Soni now Vee/Rainbow Tutu)",
     "set": {"title": "DI3 b2b: Miss Vee × Rainbow Tutu", "format": "reel", "phase": "awareness", "account": "collab"}},
    {"title": "DI3 Emmy Adelle Post", "day": "2026-10-18",
     "set": {"title": "DI3 Headliner: Emmy Adelle", "format": "reel", "phase": "awareness", "account": "collab"}},
    {"title": "DI3 TBD Post", "day": "2026-10-30",
     "set": {"title": "DI3 Sponsor spotlight + spots still open", "format": "carousel", "phase": "sponsors", "account": "di"}},
]
FIELDS = ["title", "account", "format", "phase"]


def ny_day(iso):
    """New York calendar day of a UTC timestamp (EDT until 2026-11-01, EST after)."""
    if not iso:
        return None
    t = datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(timezone.utc)
    dst_end = datetime(2026, 11, 1, 6, 0, tzinfo=timezone.utc)   # 2am EDT
    dst_start = datetime(2026, 3, 8, 7, 0, tzinfo=timezone.utc)  # 2am EST
    off = -4 if dst_start <= t < dst_end else -5
    return (t + timedelta(hours=off)).date().isoformat()


def plan(posts, today):
    live = [p for p in posts if not p.get("deleted_at")]
    changes, flags, past_di, notes = {}, [], [], []

    def setf(p, field, val, rule):
        cur = changes.setdefault(p["id"], {"post": p, "after": {}, "rules": []})
        if p.get(field) != val:
            cur["after"][field] = val
        if rule not in cur["rules"]:
            cur["rules"].append(rule)

    by_id = {p["id"]: p for p in live}
    specific_ids = set()
    for s in SPECIFIC:
        if "id" in s:
            p = by_id.get(s["id"])
            if not p or p["title"] != s["expect_title"]:
                flags.append({"title": s.get("expect_title"), "day": None, "why": "specific post not found live with that id + title - left alone"})
                continue
        else:
            m = [p for p in live if p["title"] == s["title"] and ny_day(p["scheduled_for"]) == s["day"]]
            if len(m) != 1:
                flags.append({"title": s["title"], "day": s["day"], "why": f"{len(m)} live matches by exact title + date - left alone"})
                continue
            p = m[0]
        specific_ids.add(p["id"])
        s["resolved_id"] = p["id"]

    for p in live:
        day = ny_day(p["scheduled_for"])
        is_radio = bool(RADIO_RX.search(p["title"])) or p.get("series") == "Come With Radio"
        is_di = bool(DI_RX.search(p["title"]))
        if not is_di and DI_N_RX.search(p["title"]):
            # DI2 / DI4 ... are not among the named words, so no rule matches -
            # but silently skipping a DI post is the wrong failure. Flag it.
            flags.append({"title": p["title"], "day": day, "why": "DI<number> other than DI3 - not a named DI word, left as is"})
            continue
        if is_radio and is_di:
            flags.append({"title": p["title"], "day": day, "why": "matches BOTH the radio and the DI rule - left alone"})
            continue
        if is_radio:
            setf(p, "phase", "radio", "radio")
            if p.get("account") in (None, "come_with"):
                setf(p, "account", "come_with", "radio")
            if p.get("format") is None:
                t = p["title"].lower()
                if "recap" in t:
                    setf(p, "format", "carousel", "radio")
                elif "release" in t or EP_RX.search(p["title"]):
                    setf(p, "format", "reel", "radio")
                else:
                    flags.append({"title": p["title"], "day": day, "why": "radio: no 'recap', 'release' or Ep<n> in the title - phase set, format left null"})
            if "UPCOMING EVENTS" in p["title"].upper():
                notes.append(f"'{p['title']}' ({day}) matched radio on 'CWR' in the title; it reads like an events post - check its phase.")
        if is_di:
            if day is None or day < today:
                past_di.append({"title": p["title"], "day": day})
                continue
            if p["id"] in specific_ids:
                continue  # Part 3 decides these
            setf(p, "account", "collab", "upcoming_di")
            setf(p, "phase", "awareness" if day < DI_DAY else ("event" if day == DI_DAY else "post"), "upcoming_di")

    for s in SPECIFIC:
        pid = s.get("resolved_id")
        if not pid:
            continue
        for f, v in s["set"].items():
            setf(by_id[pid], f, v, "specific")

    # Series 'Dance Infusion' with no DI word in the title: not a rule match, worth a line.
    for p in live:
        if p.get("series") == "Dance Infusion" and not DI_RX.search(p["title"]):
            notes.append(f"'{p['title']}' ({ny_day(p['scheduled_for'])}) has series Dance Infusion but no DI word in its title - not matched, left as is.")
    return live, changes, flags, past_di, notes


def main():
    src = Path(sys.argv[1])
    today = sys.argv[2] if len(sys.argv) > 2 else ny_day(datetime.now(timezone.utc).isoformat())
    posts = json.loads(src.read_text(encoding="utf-8"))[0]["posts"]
    live, changes, flags, past_di, notes = plan(posts, today)
    rows = [c for c in changes.values() if c["after"]]
    per_rule = {}
    for c in rows:
        for r in c["rules"]:
            per_rule[r] = per_rule.get(r, 0) + 1
    print(f"today (NY) = {today}; {len(posts)} rows in backup, {len(live)} live; {len(rows)} live posts change\n")
    print("BEFORE -> AFTER")
    for c in sorted(rows, key=lambda c: c["post"]["scheduled_for"] or ""):
        p = c["post"]
        diff = "; ".join(f"{f}: {p.get(f)!r} -> {v!r}" for f, v in c["after"].items())
        print(f"  {ny_day(p['scheduled_for'])}  {p['title'][:52]:52}  [{'+'.join(c['rules'])}]  {diff}")
    print("\nposts changed per rule:", per_rule)
    print("\nFLAGGED (left as is):")
    for f in flags:
        print(f"  {f['day']}  {f['title']}  - {f['why']}")
    print("\nPAST DI POSTS (unchanged):")
    for d in past_di:
        print(f"  {d['day']}  {d['title']}")
    if notes:
        print("\nNOTES:")
        for n in notes:
            print("  " + n)
    out = {"today": today, "changes": [{"id": c["post"]["id"], "set": c["after"], "rules": c["rules"]} for c in rows],
           "flags": flags, "past_di": past_di, "notes": notes, "per_rule": per_rule}
    Path(src.with_suffix(".plan.json")).write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
