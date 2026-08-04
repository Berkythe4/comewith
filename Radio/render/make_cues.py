#!/usr/bin/env python3
"""
make_cues.py — pull a station's tracklist from PROD and write a cues CSV ready
for the video render. You fill in the `start` column (when each track begins in
the recorded mix); everything else (artist, title, bpm, key) comes from the DB
so you never retype it.

Usage:
    python Radio/render/make_cues.py            # the current WORKING station
    python Radio/render/make_cues.py --station 1
    python Radio/render/make_cues.py --station 1 --out Radio/render/EP1_cues.csv

Needs SBP_PAT + SBP_REF_PROD in .env (same as the migration scripts).

Output columns:  idx,start,artist,title,bpm,song_key,camelot,duration_ms
    - start: BLANK — fill with mm:ss (e.g. 4:12) or seconds. Three ways to get
      these: Rekordbox/Engine HISTORY export (ask Claude to parse it), the
      tap-along tool (Radio/render/tap_times.html), or type them by hand.
"""
import argparse, csv, json, os, sys, urllib.request, urllib.error

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"

def env():
    e = {}
    with open(os.path.join(ROOT, ".env"), encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1); e[k] = v.strip()
    return e

def q(E, sql):
    req = urllib.request.Request(
        "https://api.supabase.com/v1/projects/%s/database/query" % E["SBP_REF_PROD"],
        data=json.dumps({"query": sql}).encode(),
        headers={"Authorization": "Bearer " + E["SBP_PAT"], "Content-Type": "application/json", "User-Agent": UA},
        method="POST")
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read().decode() or "null")
    except urllib.error.HTTPError as ex:
        raise SystemExit("HTTP %s: %s" % (ex.code, ex.read().decode()[:300]))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--station", type=int, help="station_no; default = current working station")
    ap.add_argument("--week", help="write into Radio/Week N/ instead of Radio/render/")
    ap.add_argument("--out", help="output CSV path")
    a = ap.parse_args()
    E = env()

    if a.station:
        where = "p.station_no = %d" % a.station
    else:
        where = "p.status in ('building','testing')"
    rows = q(E, """
        select p.station_no, t.sort, t.artist_name, t.title, t.bpm, t.song_key, t.camelot,
               t.duration_ms, t.show_date, t.show_venue,
               array_to_string(t.genres, '|') as genres, t.release_date
        from sc_playlist_tracks t join sc_playlists p on p.id = t.playlist_id
        where %s order by p.station_no desc, t.sort;""" % where)
    if not rows:
        raise SystemExit("No tracks found for that station.")

    station_no = rows[0]["station_no"]
    if a.out:
        out = a.out
    elif a.week:
        wk = os.path.join(ROOT, "Radio", "Week %s" % a.week); os.makedirs(wk, exist_ok=True)
        out = os.path.join(wk, "EP%s_cues.csv" % station_no)
    else:
        out = os.path.join(ROOT, "Radio", "render", "EP%s_cues.csv" % station_no)
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "duration_ms", "show_date", "show_venue", "genres", "release_date"])
        for i, r in enumerate(rows, 1):
            w.writerow([i, "", r["artist_name"] or "", r["title"] or "",
                        r["bpm"] or "", r["song_key"] or "", r["camelot"] or "", r["duration_ms"] or "",
                        r["show_date"] or "", r["show_venue"] or "",
                        r.get("genres") or "", r.get("release_date") or ""])
    print("Wrote %s  (SHOW %s, %d tracks)" % (out, station_no, len(rows)))
    print("Now fill the `start` column with mm:ss for each track, then run render_episode.py")

if __name__ == "__main__":
    main()
