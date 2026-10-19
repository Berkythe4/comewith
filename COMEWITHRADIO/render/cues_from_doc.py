#!/usr/bin/env python3
"""
cues_from_doc.py — build render cues from a typed tracklist: a .docx, .txt or .md
with one line per track in the form

    12:34 - Song Title - Artist, Artist

    python Radio/render/cues_from_doc.py \
        --doc "Radio/Elements-26/render/Elements_Ep3_Tracklist_Henry.docx" \
        --out "Radio/Elements-26/render/Elements_Ep3_Henry_cues.csv" \
        --show-venue "Elements Festival" --show-dates elements_show_dates.json \
        --merge "old_cues.csv"

Third tracklist source after Rekordbox exports and SoundCloud sets, and it shares
their sheet writer and lineup matcher rather than growing its own.

TITLE COMES FIRST here, unlike a Rekordbox export where the artist leads. That is
worth being explicit about: guessing which side of the dash is the artist from a
single line is a coin flip, and getting it backwards puts the song name in the
big display slot on every card.

--merge pulls genres, BPM and key across from an EXISTING cues file for the same
set (e.g. one built from the SoundCloud playlist), matched on title. A typed
document has the real times but rarely the metadata; the earlier file has the
metadata but guessed times. Together they are complete.
"""
import argparse, csv, io, os, re, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cues_from_rekordbox import mmss, write_times_sheet, make_show_lookup, norm_name

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

# 12:34 / 1:02:03, then " - ", then the rest. Any dash, any spacing.
LINE = re.compile(r"^\s*(\d{1,2}:\d{2}(?::\d{2})?)\s*[-–—]\s*(.+?)\s*$")


def read_doc(path):
    """Plain lines out of a .docx, or straight off a text file."""
    if path.lower().endswith(".docx"):
        xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
        paras = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
        out = []
        for p in paras:
            t = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S))
            for a, b in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'")):
                t = t.replace(a, b)
            out.append(t)
        return out
    return io.open(path, encoding="utf-8-sig").read().splitlines()


def secs(t):
    p = [int(x) for x in t.split(":")]
    return p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--times", help="also write the human sheet here")
    ap.add_argument("--ep", default="?")
    ap.add_argument("--show", default="COME WITH ELEMENTS RADIO")
    ap.add_argument("--merge", help="existing cues CSV to lift genres/bpm/key from, matched on title")
    ap.add_argument("--show-venue", default="")
    ap.add_argument("--show-dates")
    ap.add_argument("--mix", help="mix audio, for the sheet header + the last track's end")
    a = ap.parse_args()

    tracks = []
    for raw in read_doc(a.doc):
        m = LINE.match(raw)
        if not m:
            continue
        start, rest = m.group(1), m.group(2)
        # "Title - Artist". Split on the LAST dash so a title containing one
        # ("Around The World (Westend Edit) - Daft Punk, Westend") still works.
        if " - " in rest:
            title, artist = rest.rsplit(" - ", 1)
        else:
            title, artist = rest, ""
        tracks.append({"start": start, "title": title.strip(), "artist": artist.strip(),
                       "genres": "", "bpm": "", "camelot": "", "release_date": "",
                       "duration_ms": ""})
    if not tracks:
        raise SystemExit("No 'MM:SS - Title - Artist' lines found in %s" % a.doc)

    st = [secs(t["start"]) for t in tracks]
    bad = [i + 1 for i in range(len(st) - 1) if st[i + 1] <= st[i]]
    if bad:
        raise SystemExit("Times don't increase at row(s): %s" % ", ".join(map(str, bad)))
    print("%d tracks, %s -> %s" % (len(tracks), tracks[0]["start"], tracks[-1]["start"]))

    merged = 0
    if a.merge and os.path.exists(a.merge):
        old = list(csv.DictReader(open(a.merge, encoding="utf-8-sig")))
        by = {}
        for r in old:
            by.setdefault(norm_name(r.get("title", "")), r)
        for t in tracks:
            k = norm_name(t["title"])
            src = by.get(k)
            if not src:                       # containment, for "XTC" vs "XTC (feat. …)"
                for kk, r in by.items():
                    if k and (k in kk or kk in k) and abs(len(k) - len(kk)) < 26:
                        src = r; break
            if src:
                merged += 1
                for col in ("genres", "bpm", "camelot", "release_date"):
                    if not t[col] and src.get(col):
                        t[col] = src[col]
        print("merged metadata onto %d/%d from %s" % (merged, len(tracks), os.path.basename(a.merge)))

    show_date_for = make_show_lookup(a.show_dates)
    matched = 0
    with io.open(a.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "genres", "release_date", "show_date", "show_venue", "duration_ms"])
        for i, t in enumerate(tracks, 1):
            sd = show_date_for(t["artist"])
            if sd:
                matched += 1
            w.writerow([i, t["start"], t["artist"], t["title"], t["bpm"], "", t["camelot"],
                        t["genres"], t["release_date"], sd,
                        a.show_venue if (sd or a.show_venue) else "", t["duration_ms"]])
    print("wrote", a.out)
    if a.show_dates:
        print("show dates matched: %d/%d" % (matched, len(tracks)))
        for t in tracks:
            if not show_date_for(t["artist"]):
                print("   no show date: %s" % t["artist"])
    miss = [t for t in tracks if not t["genres"]]
    if miss:
        print("no genre on %d track(s):" % len(miss))
        for t in miss:
            print("   %s - %s" % (t["artist"], t["title"]))

    if a.times:
        mix = secs(tracks[-1]["start"])
        if a.mix and os.path.exists(a.mix):
            import subprocess
            mix = float(subprocess.check_output(
                ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=nk=1:nw=1", a.mix], text=True).strip())
        for i, t in enumerate(tracks):
            end = st[i + 1] if i + 1 < len(st) else mix
            t["duration_ms"] = int(max(0, end - st[i]) * 1000)
        bad_n = write_times_sheet(a.times, tracks, [t["start"] for t in tracks], a.ep, mix,
                                  "%s (typed tracklist - %d tracks)" % (os.path.basename(a.doc), len(tracks)),
                                  os.path.basename(a.mix) if a.mix else "?", show=a.show)
        print("wrote %s (non-ASCII chars: %d)" % (a.times, bad_n))


if __name__ == "__main__":
    main()
