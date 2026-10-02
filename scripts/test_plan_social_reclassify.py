"""Rule tests for scripts/plan_social_reclassify.py - no network.

    python scripts/test_plan_social_reclassify.py
"""
import copy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import plan_social_reclassify as P  # noqa: E402

fails = 0


def check(cond, label):
    global fails
    print(("PASS  " if cond else "FAIL  ") + label)
    fails += 0 if cond else 1


def post(i, title, when, **kw):
    base = {"id": f"00000000-0000-4000-8000-{i:012d}", "title": title, "scheduled_for": when, "series": None,
            "account": "come_with", "format": None, "phase": "general", "deleted_at": None,
            "caption": "c", "stage": "idea", "owner_id": "o"}
    base.update(kw)
    return base


def run(posts, today="2026-10-02"):
    specific = copy.deepcopy(P.SPECIFIC)
    P.SPECIFIC[:] = []  # rule tests run without the four named posts
    try:
        live, changes, flags, past_di, notes = P.plan(posts, today)
    finally:
        P.SPECIFIC[:] = specific
    return {c["post"]["id"]: c["after"] for c in changes.values() if c["after"]}, flags, past_di, notes


ps = [
    post(1, "CWR Ep9 Recap", "2026-11-20T23:00:00Z"),
    post(2, "CWR Ep9 Release", "2026-11-15T23:00:00Z"),
    post(3, "Come With NYC Radio EP 10", "2026-12-01T23:00:00Z"),
    post(4, "CWR Founder Video", "2026-09-13T23:00:00Z"),
    post(5, "Anything", "2026-09-13T23:00:00Z", series="Come With Radio"),
    post(6, "CWR Ep9 Release", "2026-11-15T23:00:00Z", deleted_at="2026-10-01T00:00:00Z"),
    post(7, "DI3 lineup", "2026-11-10T16:00:00Z"),
    post(8, "DI3 night of", "2026-11-14T22:00:00Z"),
    post(9, "Dance Infusion thank you", "2026-11-16T16:00:00Z"),
    post(10, "DIY merch drop", "2026-11-10T16:00:00Z"),
    post(11, "DI2 throwback", "2026-08-01T16:00:00Z"),
    post(12, "DI3 x CWR crossover", "2026-11-01T16:00:00Z"),
    post(13, "CWR Ep9 Release", "2026-11-15T23:00:00Z", account="collab", format="story"),
    post(14, "DI3 after midnight", "2026-11-15T03:30:00Z"),  # 10:30pm NY on the 14th
    post(15, "Old DI post", "2026-10-01T16:00:00Z"),
]
ch, flags, past, notes = run(ps)
I = lambda n: f"00000000-0000-4000-8000-{n:012d}"
check(ch[I(1)] == {"phase": "radio", "format": "carousel"}, "radio: recap -> carousel")
check(ch[I(2)] == {"phase": "radio", "format": "reel"}, "radio: release -> reel")
check(ch[I(3)] == {"phase": "radio", "format": "reel"}, "radio: 'EP 10' counts as Ep<number>")
check(ch[I(4)] == {"phase": "radio"} and any(f["title"] == "CWR Founder Video" for f in flags), "radio: no recap/release/Ep -> phase only, format flagged")
check(ch[I(5)]["phase"] == "radio", "radio: series Come With Radio matches without a title hit")
check(I(6) not in ch, "deleted posts are never touched")
check(ch[I(7)] == {"account": "collab", "phase": "awareness"}, "DI before 11-14 -> collab + awareness")
check(ch[I(8)] == {"account": "collab", "phase": "event"}, "DI on 11-14 -> event")
check(ch[I(9)] == {"account": "collab", "phase": "post"}, "DI after 11-14 -> post")
check(ch[I(14)]["phase"] == "event", "DI date is New York time (10:30pm on the 14th is the 14th)")
check(I(10) not in ch, "DI is a whole word: 'DIY' does not match")
check(I(11) not in ch and any(f["title"] == "DI2 throwback" for f in flags), "DI2 is not a named DI word: flagged, never guessed")
check(I(15) not in ch and any(d["title"] == "Old DI post" for d in past), "yesterday's DI post is past")
check(I(12) not in ch and any("BOTH" in f["why"] for f in flags), "a post matching radio AND DI is flagged, not guessed")
check(ch[I(13)] == {"phase": "radio"}, "radio: an existing collab account and existing format are kept (phase still set)")

live_specific, _c, sflags, _p, _n = P.plan([post(20, "DI3 TBD Post", "2026-10-30T19:05:00Z"),
                                              post(21, "DI3 TBD Post", "2026-10-30T22:00:00Z")], "2026-10-02")
check(any("2 live matches" in f["why"] for f in sflags), "a specific post with two title+date matches is flagged, not guessed")

print("\nALL PASS" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
