#!/usr/bin/env python3
"""
lvs32_cues.py — build EP 4's cues from the tracklist 32LVS sent.

His list arrived as plain text, "TIME - Title - Artist (Remix)", and it is the
cleanest of the four: every line carries a time, so unlike KRNeY's there is
nothing to align. What it does NOT always carry is an artist — five lines are
mashups or flips with no original credited, and one is an unreleased ID. Those
are recorded with the artist blank rather than guessed, and printed at the end
so Keith can fill them in if he wants them on the cards.

Release dates: the full string is searched first, so an OFFICIALLY released
remix gets its own date. When that finds nothing — which is every bootleg flip —
it falls back to the original song and the row is marked, because printing 1973
next to a SoDown remix of Pink Floyd is only right if you know that is what it
means. The report says which dates are the original's.

    python Radio/Elements-26/lvs32_cues.py            # dry, prints the report
    python Radio/Elements-26/lvs32_cues.py --write    # write the cues CSV
"""
import csv, io, os, re, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "render"))
import release_dates as RD

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

OUT = os.path.join(HERE, "render", "Elements_Ep4_32LVS_cues.csv")
SHOW_DATE, SHOW_VENUE = "2026-08-09", "Elements Festival"

# (start, title as it goes on the card, artist, base title for the fallback
#  lookup — "" means don't fall back, note)
#
# artist "" = 32LVS credited nobody and I am not inventing one.
TRACKS = [
    ("0:00",  "The Great Gig in the Sky (SoDown & Mocha Remix)", "Pink Floyd",
     "The Great Gig in the Sky", ""),
    ("1:50",  "Checkmate", "Ravenscoon", "", ""),
    ("2:30",  "Coming Up (Like That)", "Heyz", "", ""),
    ("3:40",  "Jaguar (Pierce & Hazea Remix) / Nissan Altima (32LVS Edit)", "",
     "", "mashup — no original artists credited on the list"),
    ("5:50",  "Highly / Trance Trap", "P.A.F.F.", "",
     "written as 'Highly - Trance Trap - P.A.F.F.' — read as one entry, two titles"),
    ("6:50",  "ID", "32LVS", "", "unreleased — nothing to look up"),
    ("10:10", "Purple Haze (Extrakt Flip)", "Jimi Hendrix", "Purple Haze",
     "written 'Jimmy Hendrix' — spelled Jimi on the card"),
    ("11:20", "Do U Want 2", "Buku", "", ""),
    # Written "Anyway - Hey". There is no artist "Hey" with a song called Anyway;
    # there is "Anyway" by HEYZ, on the album Who Is HEYZ — the same act as 2:30.
    ("12:20", "Anyway", "Heyz", "", "written 'Hey' — it's Heyz, same act as 2:30"),
    ("13:40", "Bring the Funk Back", "SoDown", "", ""),
    ("14:40", "Say It x Soundboy (Celo Edit) x Freeway (Aquabass Edit)", "",
     "", "three-way edit — no original artists credited"),
    ("16:00", "Eye of the Mind", "SoDown", "", ""),
    ("17:20", "Daydreaming (Effin Flip)", "", "",
     "no artist — several well-known songs share this title, so I left it open"),
    ("19:20", "Look At Me Now (Zoey808 Flip)", "", "",
     "no artist — almost certainly the Chris Brown one, but he didn't say so"),
    ("21:15", "Just Dance (Z3LLA & Lucky Flip)", "", "",
     "no artist — almost certainly Lady Gaga, but he didn't say so"),
    ("23:15", "I Feel the Earth Move (32LVS Flip)", "Carole King",
     "I Feel the Earth Move", "written 'Earth Move - Carol King' — full title, Carole spelled out"),
    ("26:50", "Fuck My Computer (PZZS Remix)", "Ninajirachi", "Fuck My Computer", ""),
    ("28:10", "TEAR U APART", "Moore Kismet", "", ""),
    ("30:28", "CSRIAC (Remix)", "Ninajirachi", "CSRIAC", ""),
    ("32:00", "Falling (Eliminate Remix)", "Frost Children", "Falling", ""),
    ("32:42", "Falling (FRAXURE Remix)", "Frost Children", "Falling",
     "second flip of the same song, 42s after the first — that's how it's written"),
    ("34:32", "thinkaboutit x Crush", "Eliminate, Pixel Terror", "thinkaboutit",
     "written 'thinkaboutit - Eliminate x Crush - Pixel Terror'; the year is "
     "Eliminate's own 'thinkaboutit', off Get Off The Internet"),
    ("36:27", "Midnight Sun x Language (Hex Cougar Remix) — 32LVS Mashup", "32LVS",
     "", "mashup — originals not credited"),
    # She styles it lowercase, and that is how it comes back from every store.
    ("40:05", "Wall of Sound (32LVS Remix)", "Charli xcx", "Wall of Sound",
     "written 'Charli XCX' — set lowercase, the way she styles it"),
]


def lookup(artist, title, base):
    """Full string first, then the original song. Returns (date, genre, whose)."""
    for q, whose in ((title, "this version"), (base, "the original")):
        if not q or not artist:
            continue
        hits = (RD.from_itunes(artist, q) or RD.from_deezer(artist, q)
                or RD.from_musicbrainz(artist, q))
        if hits:
            d, g, _src = sorted(hits)[0]
            return d, g, whose
        time.sleep(0.2)
    return "", "", ""


def main():
    write = "--write" in sys.argv
    rows = []
    print("%d tracks · looking up genre + release date…\n" % len(TRACKS))
    print("  START    ARTIST                TITLE                                    YEAR  GENRE")
    print("  " + "-" * 96)
    for i, (start, title, artist, base, note) in enumerate(TRACKS, 1):
        d, g, whose = lookup(artist, title, base)
        rows.append({"idx": i, "start": start, "artist": artist, "title": title,
                     "genres": g, "release_date": d, "whose": whose, "note": note})
        print("  %-7s  %-20s  %-40s %-5s %s%s"
              % (start, artist[:20] or "—", title[:40], (d or "—")[:4], g or "",
                 "" if whose in ("this version", "") else "  (original)"))
        time.sleep(0.2)

    print("\n  dated %d/%d · genres %d/%d"
          % (sum(1 for r in rows if r["release_date"]), len(rows),
             sum(1 for r in rows if r["genres"]), len(rows)))

    orig = [r for r in rows if r["whose"] == "the original"]
    if orig:
        print("\n  YEAR IS THE ORIGINAL SONG'S, not this remix (bootleg flips have no release date):")
        for r in orig:
            print("     %-7s %s — %s  →  %s" % (r["start"], r["artist"], r["title"][:44], r["release_date"]))

    noart = [r for r in rows if not r["artist"]]
    if noart:
        print("\n  NO ARTIST CREDITED (%d) — the card will show the title alone unless you fill these in:" % len(noart))
        for r in noart:
            print("     %-7s %s" % (r["start"], r["title"][:70]))

    notes = [r for r in rows if r["note"]]
    if notes:
        print("\n  HOW I READ THE LIST:")
        for r in notes:
            print("     %-7s %s" % (r["start"], r["note"]))

    if not write:
        print("\n(dry run — pass --write to build the cues)")
        return

    with io.open(OUT, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "genres", "release_date", "show_date", "show_venue", "duration_ms"])
        for r in rows:
            w.writerow([r["idx"], r["start"], r["artist"], r["title"], "", "", "",
                        r["genres"], r["release_date"], SHOW_DATE, SHOW_VENUE, ""])
    print("\nwrote", OUT)


if __name__ == "__main__":
    main()
