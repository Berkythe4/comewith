#!/usr/bin/env python3
"""
import_cues_to_station.py — put a rendered episode's cues into its station, so
the PUBLIC page has a tracklist.

    python scripts/import_cues_to_station.py --ep 1 \
        --cues "Radio/Elements-26/render/Elements_Ep1_cues.csv"        # dry
    python scripts/import_cues_to_station.py --ep 1 --cues ... --write

The video and the episode page read from two different places. The MP4 is
rendered from the cues CSV; the page renders `sc_playlist_tracks` through
get-station. Ep1 had 22 tracks in its video and ONE row in the database, and Ep3
had none at all — both would have published a page with an empty tracklist under
a finished mix, which is the one thing the page exists for.

Replaces the station's tracks with the cues, in cue order. Refuses on a live or
archived episode: migration 135 blocks adds there for good reason, and a
service-role script would otherwise walk straight past the trigger's intent.
"""
import argparse, csv, json, os, sys, urllib.request

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EDITION = "Come With Elements Radio"


def env():
    e = {}
    for l in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
        l = l.strip()
        if "=" in l and not l.startswith("#"):
            k, v = l.split("=", 1); e[k] = v.strip().strip('"').strip("'")
    return e


def sql(q):
    E = env()
    req = urllib.request.Request(
        "https://api.supabase.com/v1/projects/%s/database/query" % E.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk"),
        data=json.dumps({"query": q}).encode(),
        headers={"Authorization": "Bearer " + E["SBP_PAT"], "Content-Type": "application/json",
                 "User-Agent": "Mozilla/5.0 Chrome/126"}, method="POST")
    return json.loads(urllib.request.urlopen(req, timeout=90).read().decode() or "[]")


def esc(v):
    if v is None or v == "":
        return "null"
    return "'" + str(v).replace("'", "''") + "'"


def genre_array(v):
    """`genres` is text[], not text. A comma string quoted straight in is a 400,
    and the cues carry things like "Minimal / Deep Tech" that are two genres."""
    parts = [x.strip() for x in str(v or "").replace("/", ",").split(",") if x.strip()]
    if not parts:
        return "null"
    return "array[" + ",".join(esc(p) for p in parts) + "]::text[]"


def num(v):
    """bpm and duration are integers; an empty cell must be null, not ''."""
    try:
        return str(int(round(float(v))))
    except Exception:
        return "null"


def secs(t):
    """Start time in seconds, or None. A cue may legitimately have NO time — a
    track the DJ listed but never wrote a time against. Those still belong on the
    page; they just can't be placed in the running order."""
    t = str(t or "").strip()
    if not t:
        return None
    try:
        p = [int(x) for x in t.split(":")]
    except ValueError:
        return None
    return p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ep", type=int, required=True, help="edition_seq (1-4)")
    ap.add_argument("--cues", required=True)
    ap.add_argument("--edition", default=EDITION)
    # sc_playlist_tracks_source_chk allows exactly these. Passing "doc" failed the
    # whole insert with a check-constraint error, so the choices are enforced here
    # rather than discovered from a 400.
    ap.add_argument("--source", default="rekordbox",
                    choices=["soundcloud", "manual", "rekordbox", "beatport", "dj"],
                    help="how these tracks got here (see migration 102)")
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    st = sql("select id::text, name, status, station_no, "
             "(select count(*) from sc_playlist_tracks t where t.playlist_id = p.id) as have "
             "from sc_playlists p where edition_name=%s and edition_seq=%d;"
             % (esc(a.edition), a.ep))
    if not st:
        raise SystemExit("No episode %d in %s" % (a.ep, a.edition))
    st = st[0]
    if st["status"] in ("live", "archived"):
        raise SystemExit("EP %d is %s — reopen it before replacing its tracklist." % (a.ep, st["status"]))

    rows = list(csv.DictReader(open(a.cues, encoding="utf-8-sig")))
    if not rows:
        raise SystemExit("No rows in %s" % a.cues)
    print("EP %d · SHOW %s · %s" % (a.ep, st["station_no"], st["name"]))
    print("  in the database now : %s track(s)" % st["have"])
    print("  in the cues file    : %d track(s)" % len(rows))

    # Timed rows first, in time order; untimed ones keep their order after them.
    timed = sorted([r for r in rows if secs(r.get("start")) is not None],
                   key=lambda r: secs(r["start"]))
    untimed = [r for r in rows if secs(r.get("start")) is None]
    rows = timed + untimed
    if untimed:
        print("  %d track(s) have no start time — placed at the end:" % len(untimed))
        for r in untimed:
            print("     %s - %s" % (r.get("artist", ""), (r.get("title") or "")[:44]))
    # end of each track = start of the next, so the page can show real lengths
    st_s = [secs(r.get("start")) for r in rows]
    dur = []
    for i in range(len(rows)):
        if st_s[i] is None:
            dur.append(int(float(rows[i].get("duration_ms") or 0) / 1000) or 0)
        elif i + 1 < len(rows) and st_s[i + 1] is not None:
            dur.append(st_s[i + 1] - st_s[i])
        else:
            dur.append(int(float(rows[i].get("duration_ms") or 0) / 1000) or 180)

    vals = []
    for i, r in enumerate(rows):
        vals.append("(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            esc(st["id"]),
            esc("cue_%d_%03d" % (st["station_no"], i + 1)),   # synthetic id, migration 102
            esc(r.get("title")), esc(r.get("artist_name") or r.get("artist")),
            genre_array(r.get("genres")), num(r.get("bpm")), esc(r.get("camelot")),
            esc(r.get("release_date")), esc(r.get("show_date")), esc(r.get("show_venue")),
            str((i + 1) * 10), str(int(dur[i] * 1000)), esc(a.source)))
    stmt = ("begin;\n"
            "delete from sc_playlist_tracks where playlist_id = %s;\n"
            "insert into sc_playlist_tracks (playlist_id, sc_track_id, title, artist_name, genres,"
            " bpm, camelot, release_date, show_date, show_venue, sort, duration_ms, source) values\n"
            "%s;\ncommit;" % (esc(st["id"]), ",\n".join(vals)))

    for i, r in enumerate(rows[:3]):
        print("     %2d  %-7s %-24s %s" % (i + 1, r["start"], (r.get("artist") or "")[:24],
                                           (r.get("title") or "")[:40]))
    print("     …")
    if not a.write:
        print("\nDRY RUN — would replace %d row(s) with %d. Pass --write." % (st["have"], len(rows)))
        return
    sql(stmt)
    after = sql("select count(*) as n from sc_playlist_tracks where playlist_id=%s;" % esc(st["id"]))[0]["n"]
    print("\nwrote %d track(s) — station now holds %d" % (len(rows), after))


if __name__ == "__main__":
    main()
