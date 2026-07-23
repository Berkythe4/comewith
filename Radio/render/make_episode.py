#!/usr/bin/env python3
"""
make_episode.py — ONE command for the whole video. Wraps the tested pieces so you
don't run three scripts by hand.

    python Radio/render/make_episode.py --audio "Radio/Video/EP1_mix.wav"

It will, for the current working station (or --station N):
  1. pull the tracklist cues from prod (unless a filled --cues is given),
  2. get the start times — from --history <deck export> if provided, else from
     the cues if you already filled them, else stop and tell you to tap them in,
  3. render Radio/Video/EP{N}.mp4.

Options:
  --station N     specific station (default: working)
  --cues FILE     use this cues CSV as-is (skip the DB pull)
  --history FILE  a Rekordbox/Engine history export to read times from
  --cover FILE    cover image (default Radio/Artwork/Radio_Thumbnail.jpg)
  --out FILE      output mp4 (default Radio/Video/EP{N}.mp4)
  --dry           quick 1-second preview render

Nothing here writes to prod or the site.
"""
import argparse, csv, os, subprocess, sys
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
PY = sys.executable

def run(args):
    r = subprocess.run(args, cwd=ROOT)
    if r.returncode != 0:
        raise SystemExit(r.returncode)

def cues_have_times(path):
    if not os.path.exists(path):
        return False
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    return bool(rows) and all((r.get("start") or "").strip() for r in rows)

def station_no_from_cues(path):
    # EP{N}_cues.csv → N
    base = os.path.basename(path)
    import re
    m = re.search(r"EP(\d+)_cues", base)
    return m.group(1) if m else "1"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--audio", required=True)
    ap.add_argument("--station", type=int)
    ap.add_argument("--cues")
    ap.add_argument("--history")
    ap.add_argument("--cover", default=os.path.join(ROOT, "Radio", "Artwork", "Radio_Thumbnail.jpg"))
    ap.add_argument("--out")
    ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()

    if not os.path.exists(a.audio):
        raise SystemExit("Missing audio: " + a.audio)

    # 1) cues
    cues = a.cues
    if not cues:
        print("== Pulling tracklist ==")
        cmd = [PY, os.path.join(HERE, "make_cues.py")]
        if a.station:
            cmd += ["--station", str(a.station)]
        run(cmd)
        # make_cues names it EP{N}_cues.csv — find the newest one
        import glob
        found = sorted(glob.glob(os.path.join(HERE, "EP*_cues.csv")), key=os.path.getmtime)
        if not found:
            raise SystemExit("Could not find the generated cues CSV.")
        cues = found[-1]
    ep_no = station_no_from_cues(cues)
    ep_label = "EP %s" % ep_no

    # 2) times
    if a.history:
        print("== Reading times from history ==")
        run([PY, os.path.join(HERE, "import_history.py"), "--history", a.history, "--cues", cues, "--write"])
    if not cues_have_times(cues):
        print("\nThe cues still need start times before I can render:")
        print("   " + os.path.relpath(cues, ROOT).replace("\\", "/"))
        print("Fill them the easy way:")
        print("   • open Radio/render/tap_times.html, load the mix + this CSV, tap along, export")
        print("   • or pass --history <your deck's history export>")
        print("   • or type mm:ss into the `start` column")
        raise SystemExit(2)

    # 3) render
    out = a.out or os.path.join(ROOT, "Radio", "Video", "EP%s.mp4" % ep_no)
    print("== Rendering %s ==" % ep_label)
    cmd = [PY, os.path.join(HERE, "render_episode.py"),
           "--cues", cues, "--audio", a.audio, "--cover", a.cover,
           "--out", out, "--ep", ep_label]
    if a.dry:
        cmd.append("--dry")
    run(cmd)
    print("\nEpisode video ready:", os.path.relpath(out, ROOT).replace("\\", "/"))

if __name__ == "__main__":
    main()
