#!/usr/bin/env python3
"""Exit 0 if every row of a cues CSV has a start time, 1 otherwise.

Batch cannot read a CSV, and "Make Radio MP4.bat" must not hand a half-timed
tracklist to the renderer. Prints the rows that are still blank so you know
exactly what to go and fill in.
"""
import csv, io, sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

if len(sys.argv) < 2:
    sys.exit(1)

try:
    rows = list(csv.DictReader(io.open(sys.argv[1], encoding="utf-8-sig")))
except Exception as ex:
    print("   could not read the cues CSV: %s" % ex)
    sys.exit(1)

if not rows:
    print("   the cues CSV is empty")
    sys.exit(1)

blank = [r for r in rows if not (r.get("start") or "").strip()]
if blank:
    print("   %d of %d tracks have no start time:" % (len(blank), len(rows)))
    for r in blank[:8]:
        print("     %-3s %s - %s" % (r.get("idx", "?"),
                                     r.get("artist", "?"), r.get("title", "?")))
    if len(blank) > 8:
        print("     ... and %d more" % (len(blank) - 8))
    sys.exit(1)

sys.exit(0)
