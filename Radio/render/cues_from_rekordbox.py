#!/usr/bin/env python3
"""
cues_from_rekordbox.py — build the render cues from a Rekordbox playlist export.

The deck export is the truth about what was actually played and in what ORDER;
the dashboard tracklist can lag behind it. EP2 was a live example: the dashboard
had 19 tracks, the export had 23 and in a completely different order.

    python Radio/render/cues_from_rekordbox.py \
        --txt "Radio/Week 2/CWR_Ep2.txt" \
        --mix "Radio/Week 2/CWR_Ep2.073026.WAV" \
        --out "Radio/Week 2/EP2_cues.csv" \
        --times "Radio/Week 2/EP2_times.txt" --ep 2

Format notes (same ones the dashboard importer had to learn):
  * UTF-16 with a BOM, TAB separated.
  * Columns are located BY HEADER NAME, never by position — Rekordbox reorders
    them depending on the view you exported from.
  * The Artist column is sometimes EMPTY, with the artist folded into the title
    ("AC Slater - For The Funk"). Split on " - " in that case.
  * "Time" is mm:ss of the source file.

Start times are seeded as arithmetic guesses (track lengths scaled to fit the
mix), clearly labelled as such. They are a typing aid, NOT a measurement.
"""
import argparse, csv, io, os, subprocess, sys
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass


def read_rekordbox(path):
    raw = open(path, "rb").read()
    for enc in ("utf-16", "utf-16-le", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            if "\t" in text:
                break
        except Exception:
            continue
    else:
        raise SystemExit("Could not decode %s" % path)
    lines = [l for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if l.strip()]
    head = [h.strip().lstrip("﻿").strip('"') for h in lines[0].split("\t")]
    idx = {h.lower(): i for i, h in enumerate(head)}

    def col(cells, *names):
        for n in names:
            i = idx.get(n.lower())
            if i is not None and i < len(cells):
                v = cells[i].strip()
                if v:
                    return v
        return ""

    out = []
    for ln in lines[1:]:
        c = ln.split("\t")
        title = col(c, "Track Title", "Title", "Name")
        if not title:
            continue
        artist = col(c, "Artist", "Artists")
        # Empty Artist column: Rekordbox leaves it blank when the tag has the
        # artist inside the title. Split once on " - ".
        if not artist and " - " in title:
            artist, title = [p.strip() for p in title.split(" - ", 1)]
        t = col(c, "Time", "Duration", "Length")
        secs = 0
        if ":" in t:
            p = [float(x) for x in t.split(":")]
            secs = p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]
        bpm = col(c, "BPM")
        if bpm:
            try: bpm = str(int(round(float(bpm))))
            except Exception: pass
        out.append({"artist": artist, "title": title, "bpm": bpm,
                    "camelot": col(c, "Key"), "genre": col(c, "Genre"),
                    "duration_ms": int(secs * 1000)})
    return out


def mix_seconds(path):
    return float(subprocess.check_output(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=nk=1:nw=1", path], text=True).strip())


def mmss(s):
    s = int(s); return "%d:%02d" % (s // 60, s % 60)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True, help="Rekordbox playlist export")
    ap.add_argument("--mix", help="mix audio, to scale the guess times")
    ap.add_argument("--out", required=True, help="cues CSV to write")
    ap.add_argument("--times", help="also write a human fill-in sheet here")
    ap.add_argument("--ep", type=int, default=0)
    a = ap.parse_args()

    tr = read_rekordbox(a.txt)
    if not tr:
        raise SystemExit("No tracks parsed — check the export.")
    total = sum(t["duration_ms"] for t in tr) / 1000.0
    mix = mix_seconds(a.mix) if a.mix else total
    scale = (mix / total) if total else 1.0

    run, starts = 0.0, []
    for t in tr:
        starts.append(mmss(run))
        run += t["duration_ms"] / 1000.0 * scale

    with io.open(a.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot", "duration_ms"])
        for i, t in enumerate(tr, 1):
            w.writerow([i, starts[i - 1], t["artist"], t["title"], t["bpm"], "",
                        t["camelot"], t["duration_ms"]])
    print("wrote %s — %d tracks" % (a.out, len(tr)))
    print("source total %s | mix %s | overlap scale %.3f" % (mmss(total), mmss(mix), scale))

    if a.times:
        L = []
        A = L.append
        A("EP %d - COME WITH NYC RADIO : TRACK START TIMES" % a.ep)
        A("=" * 86)
        A("MIX FILE   : %s" % os.path.basename(a.mix or "?"))
        A("MIX LENGTH : %s" % mmss(mix))
        A("TRACKLIST  : %s (Rekordbox export - %d tracks, play order)" % (os.path.basename(a.txt), len(tr)))
        A("")
        A("HOW TO FILL THIS IN")
        A("  1. Type the time each track STARTS, right after the '=' sign.  Format mm:ss")
        A("  2. Within ~15 seconds is close enough - I snap it to the exact drop by audio.")
        A("  3. Track 1 is almost always 0:00.")
        A("  4. Do NOT reorder the lines. Save the file, then tell me it's done.")
        A("  5. Unsure about one? Put a ? after it, e.g.  7 = 15:25?")
        A("")
        A("The 'guess' is ARITHMETIC ONLY - track lengths scaled to fit the mix.")
        A("It is NOT from listening. A starting point, not an answer.")
        A("")
        A("-" * 86)
        for i, t in enumerate(tr, 1):
            line = "%2d = %-8s%s - %s" % (i, "", t["artist"], t["title"])
            A(line.ljust(104) + "[%s long | guess %s]" % (mmss(t["duration_ms"] / 1000.0), starts[i - 1]))
        A("-" * 86)
        txt = "\n".join(L) + "\n"
        bad = sum(1 for ch in txt if ord(ch) > 126 and ch not in "\r\n")
        io.open(a.times, "w", encoding="utf-8", newline="\r\n").write(txt)
        print("wrote %s (non-ASCII chars: %d)" % (a.times, bad))


if __name__ == "__main__":
    main()
