#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""The render cues for an episode whose cards carry EVERY show (all_shows.py).

    python render/cues_with_shows.py --txt "Episode 7/CWR_Ep7_V2.txt" \
        --times "Episode 7/EP7_times.txt" --shows "Episode 7/EP7_shows.json" \
        --station 11 --out "Episode 7/EP7_cues.csv"

  * order, bpm, key, genres  <- the Rekordbox export (what was played)
  * start times              <- the filled-in times sheet ("N = mm:ss")
  * shows / spotlight        <- preview_cards.card_for, so the video draws the
                                same cards that were reviewed as PNGs
  * release_date             <- the station's rows (track-sources / Beatport),
                                matched by position AND checked by title

Writes every column render_card reads, plus `shows` (JSON), `spotlight` and
`spotlight_line`. A blank start time stops the build rather than guessing.
"""
import argparse, csv, io, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cues_from_rekordbox import read_rekordbox
from preview_cards import card_for
from apply_station_from_rekordbox import sql, PROD

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass


def read_times(path):
    """{n: 'h:mm:ss' or 'mm:ss'} from the fill-in sheet. A '?' is kept off the value."""
    out = {}
    for ln in io.open(path, encoding="utf-8-sig"):
        m = re.match(r"^\s*(\d+)\s*=\s*(\d{1,2}(?::\d{2}){1,2})\s*\??", ln)
        if m:
            out[int(m.group(1))] = m.group(2)
    return out


def secs(v):
    p = [int(x) for x in v.split(":")]
    return p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]


def key(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True)
    ap.add_argument("--times", required=True)
    ap.add_argument("--shows", required=True)
    ap.add_argument("--station", type=int, help="station_no to take release dates from")
    ap.add_argument("--out", required=True)
    ap.add_argument("--ref", default=PROD)
    a = ap.parse_args()

    tracks = read_rekordbox(a.txt)
    shows = json.load(open(a.shows, encoding="utf-8"))
    times = read_times(a.times)
    if len(shows) != len(tracks):
        raise SystemExit("shows JSON has %d tracks, export has %d" % (len(shows), len(tracks)))
    blank = [i for i in range(1, len(tracks) + 1) if i not in times]
    if blank:
        raise SystemExit("No start time in the sheet for track(s): %s" % ", ".join(map(str, blank)))
    st = [secs(times[i]) for i in range(1, len(tracks) + 1)]
    bad = [i + 1 for i in range(1, len(st)) if st[i] <= st[i - 1]]
    if bad:
        raise SystemExit("Start times don't increase at track(s): %s" % ", ".join(map(str, bad)))

    dates = {}
    if a.station:
        rows = sql("select t.title, t.release_date::text as rd from sc_playlist_tracks t "
                   "join sc_playlists p on p.id = t.playlist_id where p.station_no = %d "
                   "order by t.sort;" % a.station, a.ref)
        if len(rows) != len(tracks):
            print("!! station %d has %d rows, export %d -- release dates skipped"
                  % (a.station, len(rows), len(tracks)))
        else:
            for i, (r, t) in enumerate(zip(rows, tracks), 1):
                # position is the join; the title check stops a reordered
                # station from dating the wrong song
                if key(r["title"])[:12] != key(t["title"])[:12]:
                    print("!! row %d title differs (%r vs %r) -- no date" % (i, r["title"], t["title"]))
                elif r["rd"]:
                    dates[i] = r["rd"]

    cols = ["idx", "start", "artist", "title", "bpm", "song_key", "camelot", "genres",
            "release_date", "show_date", "show_venue", "duration_ms",
            "shows", "spotlight", "spotlight_line"]
    with io.open(a.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for i, (t, s) in enumerate(zip(tracks, shows), 1):
            c = card_for(t, s)
            first = (c["shows"] or [{}])[0]
            w.writerow([i, times[i], t["artist"], t["title"], t["bpm"], "", t["camelot"],
                        c["genres"], dates.get(i, ""), first.get("date", ""), first.get("venue", ""),
                        t["duration_ms"], json.dumps(c["shows"], ensure_ascii=False),
                        c.get("spotlight", ""), c.get("spotlight_line", "")])
            print("%2d %8s  %-10s %-3s %s - %s" % (i, times[i], dates.get(i, "-"),
                  "DI" if c.get("spotlight") else "", t["artist"][:28], t["title"][:40]))
    print("wrote", a.out)


if __name__ == "__main__":
    main()
