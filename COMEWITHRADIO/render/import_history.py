#!/usr/bin/env python3
"""
import_history.py — read track start times out of a DJ-gear history export
(Rekordbox / Engine / a plain list) and fill the cues CSV's `start` column.

    # preview only (default — shows the mapping, writes nothing):
    python Radio/render/import_history.py --history hist.txt --cues Radio/render/EP1_cues.csv
    # commit:
    python Radio/render/import_history.py --history hist.txt --cues ... --write

Best-effort by design: it finds a time-like column, turns clock times into
offsets from the first track, maps to the cues BY ORDER, and prints the result
for you to eyeball. It never writes without --write. If it can't find times it
says so and points you at the tap tool — it won't guess.

Handled inputs:
  • Rekordbox history .txt (UTF-16 tab-separated, columns by header)
  • CSV with a time/started column
  • a plain list of mm:ss or HH:MM:SS lines (one per track, in order)
"""
import argparse, csv, io, os, re, sys
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?(?:[.,]\d+)?\s*$")

def read_text(path):
    with open(path, "rb") as f:
        b = f.read()
    if b[:2] in (b"\xff\xfe", b"\xfe\xff") or (len(b) > 4 and b[1] == 0 and b[3] == 0):
        enc = "utf-16"
    else:
        enc = "utf-8-sig"
    return b.decode(enc, errors="replace")

def to_seconds(s):
    m = TIME_RE.match(s)
    if not m:
        return None
    h = int(m.group(3) is not None and m.group(1) or 0)
    if m.group(3):  # H:M:S
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))
    return int(m.group(1)) * 60 + int(m.group(2))  # M:S

def extract_times(text):
    lines = [l for l in text.replace("\r", "").split("\n") if l.strip()]
    if not lines:
        return []
    # delimited with a header?
    delim = "\t" if "\t" in lines[0] else ("," if lines[0].count(",") >= 2 else None)
    if delim:
        rows = [l.split(delim) for l in lines]
        header = [c.strip().lower() for c in rows[0]]
        # a column whose name hints at time, OR the column that is mostly clock times
        cand = None
        for i, h in enumerate(header):
            if any(k in h for k in ("time", "start", "played", "date")):
                cand = i; break
        if cand is None:
            # scan columns for one that parses as time on most data rows
            ncol = max(len(r) for r in rows)
            best, bestn = None, 0
            for i in range(ncol):
                n = sum(1 for r in rows[1:] if i < len(r) and to_seconds(r[i].strip()) is not None)
                if n > bestn:
                    best, bestn = i, n
            if bestn >= max(2, (len(rows) - 1) // 2):
                cand = best
        if cand is not None:
            vals = []
            for r in rows[1:]:
                v = to_seconds(r[cand].strip()) if cand < len(r) else None
                vals.append(v)
            return vals
    # plain list of times, one per line
    vals = [to_seconds(l.strip()) for l in lines]
    if sum(1 for v in vals if v is not None) >= max(2, len(vals) // 2):
        return vals
    return []

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--history", required=True)
    ap.add_argument("--cues", required=True)
    ap.add_argument("--write", action="store_true")
    a = ap.parse_args()

    with open(a.cues, encoding="utf-8-sig") as f:
        cues = list(csv.DictReader(f))
        fields = cues[0].keys() if cues else []
    if not cues:
        raise SystemExit("Cues file is empty.")

    raw = extract_times(read_text(a.history))
    times = [t for t in raw if t is not None]
    if len(times) < 2:
        print("Couldn't find track times in that file.")
        print("It may not carry per-track clock times. Easiest fallback: open")
        print("Radio/render/tap_times.html, load the mix + cues, and tap along.")
        raise SystemExit(2)

    # clock times -> offsets from the first
    base = times[0]
    offsets = [max(0, t - base) for t in times]
    if offsets != sorted(offsets):
        print("Warning: extracted times aren't strictly increasing — the export may")
        print("include loaded-but-not-played entries. Review the mapping below.")

    n = min(len(cues), len(offsets))
    print("Mapping history times to %d tracks (by order):\n" % n)
    for i in range(len(cues)):
        off = offsets[i] if i < len(offsets) else None
        stamp = "%d:%02d" % (off // 60, off % 60) if off is not None else "—  (no time)"
        print("  %2d  %-10s  %s — %s" % (i + 1, stamp, (cues[i].get("artist") or "")[:24], (cues[i].get("title") or "")[:30]))
    if len(offsets) != len(cues):
        print("\n  ! history has %d times but the station has %d tracks — check the counts."
              % (len(offsets), len(cues)))

    if not a.write:
        print("\nPreview only. Re-run with --write to save these into the cues CSV.")
        return

    for i in range(len(cues)):
        if i < len(offsets):
            cues[i]["start"] = "%d:%02d" % (offsets[i] // 60, offsets[i] % 60)
    with open(a.cues, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(fields))
        w.writeheader(); w.writerows(cues)
    print("\nWrote start times into", a.cues)

if __name__ == "__main__":
    main()
