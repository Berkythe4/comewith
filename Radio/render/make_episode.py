#!/usr/bin/env python3
"""
make_episode.py — ONE command for the whole video. Wraps the tested pieces so you
don't run three scripts by hand.

    python Radio/render/make_episode.py --week 2   # finds the mix + cues in Radio/Week 2/

It will, for the current working station (or --station N):
  1. pull the tracklist cues from prod (unless a filled --cues is given),
  2. get the start times — from --history <deck export> if provided, else from
     the cues if you already filled them, else stop and tell you to tap them in,
  3. render Radio/Week {N}/CWR_Ep{N}_YouTube.mp4.

Options:
  --station N     specific station (default: working)
  --cues FILE     use this cues CSV as-is (skip the DB pull)
  --history FILE  a Rekordbox/Engine history export to read times from
  --cover FILE    cover image (default Radio/Artwork/Radio_Thumbnail.jpg)
  --out FILE      output mp4 (default Radio/Week {N}/CWR_Ep{N}_YouTube.mp4)
  --dry           quick 1-second preview render

Nothing here writes to prod or the site.
"""
import argparse, csv, os, subprocess, sys
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
PY = sys.executable
sys.path.insert(0, HERE)
from _paths import episode_dir

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

def resolve_episode(ep_no, edition=None):
    """Episode number -> the row in sc_playlists that IS that episode.

    THESE ARE TWO DIFFERENT NUMBERS and conflating them ships a wrong video.
    `station_no` is the global show counter; the episode number is that show's
    place in ITS OWN series. The Elements run took shows 3-6, so NYC Radio Ep3
    is SHOW 7 -- and Ep3's own number lives only in the name, because
    edition_seq is NULL for the flagship show.

    Returns (station_no, ep_label, mix_by, drop_date, next_drop) or None.
    """
    import re as _re
    sys.path.insert(0, HERE)
    from make_cues import env as _env, q as _q
    E = _env()
    rows = _q(E, """
        select station_no, name, edition_name, edition_seq, mix_by,
               drop_date::text as drop_date,
               (select drop_date::text from sc_playlists n
                where n.drop_date > p.drop_date order by n.drop_date limit 1) as next_drop
        from sc_playlists p order by station_no;""")

    def ep_of(r):
        if r.get("edition_seq"):
            return int(r["edition_seq"])
        m = _re.search(r"Ep\s*0*(\d+)", r.get("name") or "", _re.I)
        return int(m.group(1)) if m else None

    if edition:
        cand = [r for r in rows if (r.get("edition_name") or "") == edition]
    else:
        cand = [r for r in rows if not r.get("edition_name")]
    hits = [r for r in cand if ep_of(r) == int(ep_no)]
    if len(hits) != 1:
        return None
    r = hits[0]
    return (r["station_no"], "EP %d" % int(ep_no), r.get("mix_by") or "",
            r.get("drop_date") or "", r.get("next_drop") or "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", "--week", dest="week", help="episode number / folder — auto-finds the mix, history, cues in Radio/Week N/ and writes CWR_EpN_YouTube.mp4 there")
    ap.add_argument("--audio")
    ap.add_argument("--station", type=int)
    ap.add_argument("--cues")
    ap.add_argument("--history")
    ap.add_argument("--cover", default=os.path.join(ROOT, "Radio", "Artwork", "Radio_Thumbnail.jpg"))
    ap.add_argument("--out")
    ap.add_argument("--next-date", dest="next_date_in",
                    help="the NEXT episode's drop date, YYYY-MM-DD. Beats whatever "
                         "the database has scheduled -- that is a plan, this is the "
                         "date going on the closing slide.")
    ap.add_argument("--ep", help="episode number (its place in its own series) if not using --week")
    ap.add_argument("--title", default="Come With NYC Radio",
                    help="the SHOW name drawn in the header. 'Come With Radio' is the brand, "
                         "not the show — the NYC show is 'Come With NYC Radio'.")
    ap.add_argument("--edition-name", dest="edition_name",
                    help="for a special edition, e.g. 'Come With Elements Radio'")
    ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()

    # --week: a whole episode lives in Radio/Week N/. Auto-discover its files so
    # the whole render is just `make_episode.py --week 1`.
    wk = None
    if a.week:
        wk = episode_dir(ROOT, a.week)
        if not wk:
            raise SystemExit("No folder for episode %s (looked for Radio/Episode %s)" % (a.week, a.week))
        a.audio = a.audio or _find(wk, ["*.wav", "*.aiff", "*.flac", "*.mp3", "*.m4a"])
        a.history = a.history or _find(wk, ["*.m3u8", "HISTORY*.txt", "*history*.txt", "*.cue"])
        a.cues = a.cues or _find(wk, ["EP*_cues.csv", "*_cues.csv"])
        a.out = a.out or os.path.join(wk, "CWR_Ep%s_YouTube.mp4" % a.week)

    if not a.audio or not os.path.exists(a.audio):
        raise SystemExit("Missing audio (looked in %s): %s" % (wk or "--audio", a.audio))

    # 0) which episode is this, and which SHOW is it? (see resolve_episode)
    ep_label = mixed_by = drop_date = next_date = ""
    ep_num = a.week or a.ep
    if ep_num:
        got = resolve_episode(ep_num, a.edition_name)
        if not got:
            raise SystemExit(
                "Could not find episode %s in sc_playlists. Pass --station N "
                "(the show number) explicitly, or check the episode's name in "
                "the dashboard reads like 'Come With NYC Radio Ep%s'."
                % (ep_num, ep_num))
        station_no, ep_label, mixed_by, drop_date, next_date = got
        if a.next_date_in:
            next_date = a.next_date_in
        a.station = a.station or station_no
        print("== %s = SHOW %s -- %s, drops %s, next %s ==" % (
            ep_label, station_no, mixed_by or "?", drop_date or "?", next_date or "?"))

    # 1) cues
    cues = a.cues
    if not cues:
        print("== Pulling tracklist ==")
        cmd = [PY, os.path.join(HERE, "make_cues.py")]
        if a.station:
            cmd += ["--station", str(a.station)]
        if wk:
            cmd += ["--out", os.path.join(wk, "EP%s_cues.csv" % (a.week or a.station))]
        run(cmd)
        search = wk or HERE
        found = sorted(_glob.glob(os.path.join(search, "EP*_cues.csv")), key=os.path.getmtime)
        if not found:
            raise SystemExit("Could not find the generated cues CSV.")
        cues = found[-1]
    # The on-screen label was resolved from the DB above (episode number, not
    # station_no). Fall back to the cues filename only when nothing resolved.
    if not ep_label:
        ep_label = "EP %s" % station_no_from_cues(cues)

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

    # 3) render. mixed_by / drop_date / next_date came from resolve_episode.

    epn = a.week or ep_label.split()[-1]
    out = a.out or os.path.join(episode_dir(ROOT, epn, must_exist=False), "CWR_Ep%s_YouTube.mp4" % epn)
    print("== Rendering %s ==" % ep_label)
    cmd = [PY, os.path.join(HERE, "render_episode.py"),
           "--audio", a.audio, "--cover", a.cover, "--out", out, "--ep", ep_label,
           "--title", a.title, "--abitrate", "320k",
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
