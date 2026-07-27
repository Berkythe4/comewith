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

def _json_has_times(path):
    try:
        import json
        d = json.load(open(path, encoding="utf-8"))
        return bool(d) and all(r.get("start_sec") not in (None, "") for r in d)
    except Exception:
        return False

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

import glob as _glob
def _find(folder, patterns):
    for p in patterns:
        hits = sorted(_glob.glob(os.path.join(folder, p)), key=os.path.getmtime)
        if hits:
            return hits[-1]
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--week", help="episode week folder — auto-finds the mix, history, cues in Radio/Week N/ and writes CWR_EpN_YouTube.mp4 there")
    ap.add_argument("--audio")
    ap.add_argument("--station", type=int)
    ap.add_argument("--cues")
    ap.add_argument("--history")
    ap.add_argument("--cover", default=os.path.join(ROOT, "Radio", "Artwork", "Radio_Thumbnail.jpg"))
    ap.add_argument("--out")
    ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()

    # --week: a whole episode lives in Radio/Week N/. Auto-discover its files so
    # the whole render is just `make_episode.py --week 1`.
    wk = None
    if a.week:
        wk = os.path.join(ROOT, "Radio", "Week %s" % a.week)
        if not os.path.isdir(wk):
            raise SystemExit("No folder: " + wk)
        a.audio = a.audio or _find(wk, ["*.wav", "*.aiff", "*.flac", "*.mp3", "*.m4a"])
        a.history = a.history or _find(wk, ["*.m3u8", "HISTORY*.txt", "*history*.txt", "*.cue"])
        a.cues = a.cues or _find(wk, ["EP*_cues.csv", "*_cues.csv"])
        a.out = a.out or os.path.join(wk, "CWR_Ep%s_YouTube.mp4" % a.week)
        a.station = a.station or (int(a.week) if str(a.week).isdigit() else None)

    if not a.audio or not os.path.exists(a.audio):
        raise SystemExit("Missing audio (looked in %s): %s" % (wk or "--audio", a.audio))

    # 1) cues
    cues = a.cues
    if not cues:
        print("== Pulling tracklist ==")
        cmd = [PY, os.path.join(HERE, "make_cues.py")]
        if a.station:
            cmd += ["--station", str(a.station)]
        if wk:
            cmd += ["--out", os.path.join(wk, "EP%s_cues.csv" % (a.station or a.week))]
        run(cmd)
        search = wk or HERE
        found = sorted(_glob.glob(os.path.join(search, "EP*_cues.csv")), key=os.path.getmtime)
        if not found:
            raise SystemExit("Could not find the generated cues CSV.")
        cues = found[-1]
    ep_no = station_no_from_cues(cues)
    ep_label = "EP %s" % ep_no

    # 2) times. A hand-verified tracklist.json (from match_mix + your corrections)
    # is the source of truth — if it's there with times, use it and skip the
    # history/cues dance entirely.
    tl_json = _find(wk, ["tracklist.json"]) if wk else None
    if not (tl_json and _json_has_times(tl_json)):
        tl_json = None
        if a.history:
            print("== Reading times from history ==")
            try:
                run([PY, os.path.join(HERE, "import_history.py"), "--history", a.history, "--cues", cues, "--write"])
            except SystemExit:
                print("(history had no usable times — will need tap-along or a tracklist.json)")
        if not cues_have_times(cues):
            print("\nNo start times yet. Get them one of these ways, then re-run:")
            print("   • audio-match:  python Radio/render/match_mix.py --mix <wav> --tracks <folder> --tracklist %s" % os.path.relpath(cues, ROOT).replace("\\", "/"))
            print("   • tap-along:    open Radio/render/tap_times.html, load the mix + this CSV, export")
            print("   • or type mm:ss into the `start` column of the cues CSV")
            raise SystemExit(2)

    # 3) render — pull the intro/closing meta (who mixed it, this drop, next drop)
    # from prod so the bookend slides fill themselves in.
    mixed_by = drop_date = next_date = ""
    try:
        sys.path.insert(0, HERE)
        from make_cues import env as _env, q as _q
        E = _env()
        meta = _q(E, """
            select mix_by, drop_date::text,
                   (select drop_date::text from sc_playlists n
                    where n.drop_date > p.drop_date order by n.drop_date limit 1) as next_drop
            from sc_playlists p where p.station_no = %s limit 1;""" % ep_no)
        if meta:
            mixed_by = meta[0].get("mix_by") or ""
            drop_date = meta[0].get("drop_date") or ""
            next_date = meta[0].get("next_drop") or ""
    except Exception as ex:
        print("(couldn't read episode meta for the intro/closing — rendering with what we have: %s)" % ex)

    out = a.out or os.path.join(ROOT, "Radio", "Week %s" % ep_no, "CWR_Ep%s_YouTube.mp4" % ep_no)
    print("== Rendering %s ==" % ep_label)
    cmd = [PY, os.path.join(HERE, "render_episode.py"),
           "--audio", a.audio, "--cover", a.cover, "--out", out, "--ep", ep_label, "--abitrate", "320k",
           "--mixed-by", mixed_by, "--drop-date", drop_date, "--next-date", next_date]
    if tl_json:
        cmd += ["--json", tl_json, "--meta", cues]
    else:
        cmd += ["--cues", cues]
    if a.dry:
        cmd.append("--dry")
    run(cmd)
    print("\nEpisode video ready:", os.path.relpath(out, ROOT).replace("\\", "/"))

if __name__ == "__main__":
    main()
