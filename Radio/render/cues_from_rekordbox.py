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
import argparse, csv, io, json, os, re, subprocess, sys
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
        # "Timestamp Track Title" is its own Rekordbox column — the one that
        # prepends the cue time. Another reason columns are found by NAME: this
        # export has no "Track Title" at all, so a positional reader sees nothing.
        title = col(c, "Timestamp Track Title", "Track Title", "Title", "Name")
        if not title:
            continue
        # A cue time typed into the title ("00:00 - Appetite", "1:01:25 - Shake
        # Something", and sometimes "04:00-  Demon Time" with no space). These are
        # REAL start times someone wrote down, so they beat the arithmetic guess —
        # but only when asked for, since a song can legitimately open with a
        # number. Tolerates h:mm:ss, mm:ss, and any spacing around the dash.
        cue = None
        m = re.match(r"^\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*[-–—]\s*(.+)$", title)
        if m:
            p = [int(x) for x in m.group(1).split(":")]
            cue = p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]
            title = m.group(2).strip()
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
        # Strip stray U+FEFF. Rekordbox sometimes leaves a byte-order mark INSIDE
        # a field ("I've Been Waiting﻿﻿ (Extended)"). It is an encoding
        # artifact, not part of the name — everything else about the title, symbols
        # included, is left exactly as written.
        title = title.replace("﻿", "")
        artist = artist.replace("﻿", "")
        out.append({"artist": artist, "title": title, "bpm": bpm,
                    "camelot": col(c, "Key"), "genre": col(c, "Genre"),
                    "cue": cue, "duration_ms": int(secs * 1000)})
    return out


def mix_seconds(path):
    return float(subprocess.check_output(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=nk=1:nw=1", path], text=True).strip())


def mmss(s):
    s = int(s); return "%d:%02d" % (s // 60, s % 60)


def norm_name(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def artist_candidates(artist):
    """Every name a credit might be filed under, most specific first.

    A credit is rarely just the booked act: "Matroda featuring Dances With White
    Girls", "Opiuo and Wreckno", "Nikita, the Wicked". Punctuation and the
    feature tail are what stop a lineup lookup from landing, so both go.
    """
    a = (artist or "").strip()
    lead = re.sub(r"\b(feat|ft|featuring|with|presents|pres)\b.*$", " ", a, flags=re.I)
    parts = re.split(r"[,&/+]| and | x | vs | versus ", lead, flags=re.I)
    return [p.strip() for p in ([a, lead] + parts) if p and p.strip()]


def make_show_lookup(path):
    """artist -> 'YYYY-MM-DD' from a {name: date} JSON, matched forgivingly."""
    if not path:
        return lambda _a: ""
    shows = {norm_name(k): v for k, v in json.load(open(path, encoding="utf-8")).items()}

    def lookup(artist):
        for nm in artist_candidates(artist):
            v = shows.get(norm_name(nm))
            if v:
                return v
        return ""
    return lookup


def write_times_sheet(path, tracks, starts, ep, mix_secs, source_line, mix_name="?", show="COME WITH NYC RADIO"):
    """The human fill-in sheet. Shared so a second tracklist source (a SoundCloud
    playlist, say) produces the SAME document rather than a near-copy of it.

    Deliberately plain ASCII and CRLF: it gets opened in Notepad on a phone or a
    laptop and typed into, and a smart quote or a lone LF makes that worse.
    """
    L = []
    A = L.append
    A("%s %s - %s : TRACK START TIMES" % ("EP", ep, show))
    A("=" * 86)
    A("MIX FILE   : %s" % mix_name)
    A("MIX LENGTH : %s" % mmss(mix_secs))
    A("TRACKLIST  : %s" % source_line)
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
    for i, t in enumerate(tracks, 1):
        line = "%2d = %-8s%s - %s" % (i, "", t["artist"], t["title"])
        A(line.ljust(104) + "[%s long | guess %s]" % (mmss(t["duration_ms"] / 1000.0), starts[i - 1]))
    A("-" * 86)
    txt = "\n".join(L) + "\n"
    io.open(path, "w", encoding="utf-8", newline="\r\n").write(txt)
    return sum(1 for ch in txt if ord(ch) > 126 and ch not in "\r\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True, help="Rekordbox playlist export")
    ap.add_argument("--mix", help="mix audio, to scale the guess times")
    ap.add_argument("--out", required=True, help="cues CSV to write")
    ap.add_argument("--times", help="also write a human fill-in sheet here")
    ap.add_argument("--ep", type=int, default=0)
    ap.add_argument("--times-in-title", action="store_true",
                    help="the export already carries real start times in the title "
                         "('00:00 - Appetite') — use those instead of guessing")
    ap.add_argument("--show-venue", default="",
                    help="venue chip for every card, e.g. 'Elements Festival'")
    ap.add_argument("--show-dates", help="JSON {artist: 'YYYY-MM-DD'} for the date chip")
    a = ap.parse_args()

    tr = read_rekordbox(a.txt)
    if not tr:
        raise SystemExit("No tracks parsed — check the export.")
    total = sum(t["duration_ms"] for t in tr) / 1000.0
    mix = mix_seconds(a.mix) if a.mix else total
    scale = (mix / total) if total else 1.0

    if a.times_in_title:
        missing = [i for i, t in enumerate(tr, 1) if t.get("cue") is None]
        if missing:
            raise SystemExit("--times-in-title, but these rows have no time in the title: %s"
                             % ", ".join(map(str, missing)))
        starts = [mmss(t["cue"]) for t in tr]
        drift = [i for i in range(1, len(tr)) if tr[i]["cue"] <= tr[i - 1]["cue"]]
        if drift:
            raise SystemExit("Times in the title aren't increasing at row(s): %s"
                             % ", ".join(str(i + 1) for i in drift))
    else:
        run, starts = 0.0, []
        for t in tr:
            starts.append(mmss(run))
            run += t["duration_ms"] / 1000.0 * scale

    # Per-artist show chips. The date is matched on the FIRST credited artist —
    # a Rekordbox artist field is often "Romeo, Biscits" while the lineup lists
    # one of them — and falls back to any credited name that we know.
    show_date_for = make_show_lookup(a.show_dates)

    # Write EVERY column render_episode reads. A cues file missing `genres`,
    # `release_date`, `show_date` or `show_venue` doesn't error — that part of the
    # card just isn't drawn, silently (see NOTES_WEEKLY_RELEASE.md).
    matched = 0
    with io.open(a.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "genres", "release_date", "show_date", "show_venue", "duration_ms"])
        for i, t in enumerate(tr, 1):
            sd = show_date_for(t["artist"])
            if sd:
                matched += 1
            w.writerow([i, starts[i - 1], t["artist"], t["title"], t["bpm"], "",
                        t["camelot"], t.get("genre", ""), "", sd,
                        a.show_venue if (sd or a.show_venue) else "", t["duration_ms"]])
    print("wrote %s — %d tracks" % (a.out, len(tr)))
    if a.times_in_title:
        print("start times: read from the export (real cues), last starts %s" % starts[-1])
    else:
        print("source total %s | mix %s | overlap scale %.3f" % (mmss(total), mmss(mix), scale))
    if a.show_dates:
        print("show dates matched: %d/%d" % (matched, len(tr)))
        for t in tr:
            if not show_date_for(t["artist"]):
                print("   no show date: %s" % t["artist"])

    if a.times:
        bad = write_times_sheet(
            a.times, tr, starts, a.ep, mix,
            "%s (Rekordbox export - %d tracks, play order)" % (os.path.basename(a.txt), len(tr)),
            os.path.basename(a.mix or "?"))
        print("wrote %s (non-ASCII chars: %d)" % (a.times, bad))


if __name__ == "__main__":
    main()
