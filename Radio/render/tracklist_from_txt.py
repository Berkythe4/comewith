#!/usr/bin/env python3
"""
tracklist_from_txt.py — the typed tracklist is the source of truth.

    python Radio/render/tracklist_from_txt.py --week 3            # show me what you read
    python Radio/render/tracklist_from_txt.py --week 3 --out-cues # write the render cues
    python Radio/render/tracklist_from_txt.py --week 3 --write-order  # fix the site's order

Reads `Radio/Week N/Track List *.txt` — the file written by hand after the set,
in the form the hand actually writes it:

    1 John Summit, Chainsmokers, Ilsey - ALL THE TIME - 0:00
    11 Twisted - Budino 22:23            <- dash before the time is optional
    24 On Again - Mau P -47:02           <- so is the space

That one file carries BOTH things nothing else has: the ORDER the set was
actually played in, and the START TIME of every track. The database has
everything else (SoundCloud links, artwork, durations, the artists' upcoming NYC
shows), so the two are joined rather than either being retyped.

WHY THIS EXISTS: the dashboard tracklist is what was *planned*. Nothing updates
it to what was *played* unless someone does it by hand, and on Ep 3 nobody did —
the site still had the running order from when a different person was going to
mix it. Same songs, wrong order. A video built from that order would have named
the wrong track on every card.

Matching is on title first, then artist, ignoring case, punctuation, remix
brackets and which side of the dash each one landed on — because the typed line
("I Got 5 On It - LNZ (Franky Rizardo Remix)") and the database row
("LNZ - I Got 5 On It (Franky Rizardo Remix)") describe the same track in
different word order. Anything it cannot match confidently is REPORTED, never
guessed.
"""
import argparse, csv, glob, io, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
from _paths import episode_dir

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

# "12" then the words then "4:56" or "1:02:03" at the very end. The dash and the
# space before the time are both optional because both go missing in real files.
LINE = re.compile(r"^\s*(\d{1,3})[\.\)]?\s+(.*?)[\s\-–—]*(\d{1,2}:\d{2}(?::\d{2})?)\s*$")

NOISE = re.compile(r"\((original mix|extended mix|radio edit)\)", re.I)


def norm(s):
    """Down to comparable letters: no case, no punctuation, no filler words."""
    s = (s or "").lower()
    s = NOISE.sub(" ", s)
    s = re.sub(r"[\(\)\[\]]", " ", s)
    s = re.sub(r"\b(feat|ft|featuring|remix|edit|vip|bootleg|instrumental|mix)\b", " ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def secs(t):
    p = [int(x) for x in t.split(":")]
    return p[0] * 3600 + p[1] * 60 + p[2] if len(p) == 3 else p[0] * 60 + p[1]


def mmss(n):
    n = int(round(n))
    return "%d:%02d" % (n // 60, n % 60)


def parse(path):
    """-> [{'n':1, 'raw':'...', 'a':'...', 'b':'...', 'start':0}] in file order."""
    out = []
    for line in io.open(path, encoding="utf-8-sig", errors="replace"):
        m = LINE.match(line.rstrip())
        if not m:
            continue
        n, body, t = m.group(1), m.group(2).strip(), m.group(3)
        # Split on the LAST dash: "Hide & Seek - Stormzy (Cataract Edit)" keeps
        # the ampersand-y title intact, and "A - B - C" puts C on the right.
        parts = re.split(r"\s+[-–—]\s+", body)
        if len(parts) >= 2:
            a, b = " - ".join(parts[:-1]), parts[-1]
        else:
            a, b = body, ""
        out.append({"n": int(n), "raw": body, "a": a, "b": b, "start": secs(t)})
    return out


def db_tracks(station_no):
    from make_cues import env, q
    return q(env(), """
        select t.id, t.sort, t.artist_name, t.title, t.duration_ms,
               t.permalink_url, t.genres, t.release_date,
               t.show_date::text as show_date, t.show_venue, t.show_cost, t.show_url,
               t.bpm, t.song_key, t.camelot
        from sc_playlists p join sc_playlist_tracks t on t.playlist_id = p.id
        where p.station_no = %d order by t.sort;""" % int(station_no))


def sim(a, b):
    """Jaccard over words. Word ORDER does not matter -- the typed line and the
    database row routinely disagree about it -- but the WORDS do, which is the
    whole point: an artist's two tracks share the artist and nothing else."""
    A, B = set((a or "").split()), set((b or "").split())
    if not A or not B:
        return 0.0
    return len(A & B) / float(len(A | B))


def match(entries, rows):
    """Pair typed lines with DB rows. Returns (pairs, unmatched, unused).

    Scored on the WHOLE line against the whole row, then assigned best-pair-first
    across the entire set -- not one entry at a time.

    Both of those are load-bearing. Scoring the artist separately let a perfect
    artist hit outrank the title, so Baauer's two tracks, Franky Rizardo's two
    and Casey Club's two all silently swapped. Assigning per-entry in file order
    let an early weak match consume the row a later strong one needed.
    """
    grid = []
    for i, e in enumerate(entries):
        typed = norm("%s %s" % (e["a"], e["b"])) or norm(e["raw"])
        for j, r in enumerate(rows):
            row = norm("%s %s" % (r.get("artist_name") or "", r.get("title") or ""))
            grid.append((sim(typed, row), i, j))
    grid.sort(reverse=True)

    took_e, took_r, pairs = {}, set(), []
    for sc, i, j in grid:
        if sc < 0.30 or i in took_e or j in took_r:
            continue
        took_e[i] = (rows[j], sc)
        took_r.add(j)
    for i, e in enumerate(entries):
        if i in took_e:
            r, sc = took_e[i]
            pairs.append((e, r, int(round(sc * 100))))
    unmatched = [e for i, e in enumerate(entries) if i not in took_e]
    unused = [r for j, r in enumerate(rows) if j not in took_r]
    return pairs, unmatched, unused


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", "--week", dest="week", help="Radio/Episode N — finds the tracklist txt inside")
    ap.add_argument("--txt", help="the typed tracklist, if not in the week folder")
    ap.add_argument("--station", type=int, help="show number; default resolved from --week")
    ap.add_argument("--out-cues", nargs="?", const=True,
                    help="write the render cues CSV (default: Radio/Week N/EPN_cues.csv)")
    ap.add_argument("--write-order", action="store_true",
                    help="UPDATES THE LIVE SITE: renumbers sort to the played order")
    a = ap.parse_args()

    wk = episode_dir(ROOT, a.week) if a.week else None
    txt = a.txt
    if not txt and wk:
        hits = sorted(glob.glob(os.path.join(wk, "Track List*.txt")) +
                      glob.glob(os.path.join(wk, "*racklist*.txt")))
        txt = hits[0] if hits else None
    if not txt or not os.path.exists(txt):
        raise SystemExit("No tracklist txt found (looked in %s)" % (wk or "--txt"))

    station = a.station
    if not station and a.week:
        import make_episode
        got = make_episode.resolve_episode(a.week)
        if not got:
            raise SystemExit("Could not resolve episode %s to a show number." % a.week)
        station = got[0]

    entries = parse(txt)
    print("Read %d tracks from %s" % (len(entries), os.path.basename(txt)))
    if not entries:
        raise SystemExit("Parsed nothing — check the file has 'N Title - Artist - m:ss' lines.")

    # times must climb; a typo here silently reorders the video
    bad = [(entries[i - 1], entries[i]) for i in range(1, len(entries))
           if entries[i]["start"] <= entries[i - 1]["start"]]
    for p, c in bad:
        print("  !! time goes backwards: #%d %s (%s) then #%d %s (%s)"
              % (p["n"], p["raw"][:34], mmss(p["start"]), c["n"], c["raw"][:34], mmss(c["start"])))

    rows = db_tracks(station)
    print("Show %s has %d tracks in the dashboard\n" % (station, len(rows)))
    pairs, unmatched, unused = match(entries, rows)

    for e, r, sc in pairs:
        flag = "   " if sc >= 90 else " ~ "
        print("%s%2d  %-6s  %-26s %s" % (flag, e["n"], mmss(e["start"]),
                                         (r["artist_name"] or "")[:26], (r["title"] or "")[:40]))
    if unmatched:
        print("\n!! %d typed track(s) matched NOTHING in the dashboard:" % len(unmatched))
        for e in unmatched:
            print("     #%d  %s" % (e["n"], e["raw"]))
    if unused:
        print("\n!! %d dashboard track(s) were NOT played:" % len(unused))
        for r in unused:
            print("     sort %-5s %s - %s" % (r["sort"], r["artist_name"], r["title"]))
    if unmatched or unused:
        print("\n   Fix those in the dashboard before writing anything.")

    if a.out_cues:
        out = a.out_cues if isinstance(a.out_cues, str) else \
            os.path.join(wk, "EP%s_cues.csv" % a.week)
        cols = ["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                "duration_ms", "genres", "release_date", "show_date", "show_venue",
                "show_cost", "show_url"]
        with io.open(out, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for i, (e, r, _) in enumerate(pairs, 1):
                g = r.get("genres")
                w.writerow({"idx": i, "start": mmss(e["start"]),
                            "artist": r["artist_name"], "title": r["title"],
                            "bpm": r.get("bpm") or "", "song_key": r.get("song_key") or "",
                            "camelot": r.get("camelot") or "",
                            "duration_ms": r.get("duration_ms") or "",
                            "genres": ", ".join(g) if isinstance(g, list) else (g or ""),
                            "release_date": r.get("release_date") or "",
                            "show_date": r.get("show_date") or "",
                            "show_venue": r.get("show_venue") or "",
                            "show_cost": r.get("show_cost") or "",
                            "show_url": r.get("show_url") or ""})
        print("\nWrote %s  (%d tracks)" % (os.path.relpath(out, ROOT).replace("\\", "/"), len(pairs)))

    if a.write_order:
        if unmatched or unused:
            raise SystemExit("\nRefusing to touch the live site while anything is unmatched.")
        from make_cues import env, q
        E = env()
        # Two passes: park the rows in a range nothing else uses, then set the
        # real order. A single pass collides whenever the new sort of one row is
        # the current sort of another, which is most of them.
        for i, (_, r, _) in enumerate(pairs, 1):
            q(E, "update sc_playlist_tracks set sort = %d where id = '%s';" % (900000 + i, r["id"]))
        for i, (_, r, _) in enumerate(pairs, 1):
            q(E, "update sc_playlist_tracks set sort = %d where id = '%s';" % (i * 10, r["id"]))
        print("\nLIVE SITE UPDATED: show %s reordered to the played order (%d tracks)."
              % (station, len(pairs)))


if __name__ == "__main__":
    main()
