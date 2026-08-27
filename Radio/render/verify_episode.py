#!/usr/bin/env python3
"""
verify_episode.py — the checks from NOTES_WEEKLY_RELEASE.md §9, run for you.

    python Radio/render/verify_episode.py --week 3

Never trust the render because it printed DONE. This ffprobes the finished MP4
against the mix it came from and pulls real frames to look at, because the two
worst bugs this project has shipped were both invisible to the console:

  * a 65-minute video whose new closing line silently did not draw
  * an mp4 with a plausible size and duration whose audio was -91 dB

Exit code 0 = every check passed. Non-zero = look at what it says before you
upload anything.
"""
import argparse, glob, json, os, subprocess, sys
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
from _paths import episode_dir

OK, BAD = "  ok  ", " FAIL "


def probe(path):
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-print_format", "json",
         "-show_format", "-show_streams", path],
        capture_output=True, text=True)
    if out.returncode != 0:
        raise SystemExit("ffprobe could not read %s\n%s" % (path, out.stderr.strip()))
    return json.loads(out.stdout)


def stream(info, kind):
    for s in info.get("streams", []):
        if s.get("codec_type") == kind:
            return s
    return {}


def clock(sec):
    sec = int(round(sec))
    return "%d:%02d:%02d" % (sec // 3600, sec % 3600 // 60, sec % 60)


def find(folder, patterns):
    for p in patterns:
        hits = sorted(glob.glob(os.path.join(folder, p)), key=os.path.getmtime)
        if hits:
            return hits[-1]
    return None


def mean_volume(path):
    """A file can have perfect metadata and silence in it. EP 4 of the Elements
    run produced exactly that when two ffmpeg jobs wrote at once."""
    r = subprocess.run(["ffmpeg", "-i", path, "-af", "volumedetect",
                        "-f", "null", "-"], capture_output=True, text=True)
    for line in r.stderr.splitlines():
        if "mean_volume:" in line:
            try:
                return float(line.split("mean_volume:")[1].split("dB")[0])
            except Exception:
                return None
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", "--week", dest="week", help="episode number — Radio/Episode N")
    ap.add_argument("--mp4", help="the finished video (default: found in the week folder)")
    ap.add_argument("--audio", help="the mix it came from (default: found in the week folder)")
    a = ap.parse_args()

    wk = None
    if a.week:
        wk = episode_dir(ROOT, a.week) or os.path.join(ROOT, "Radio", "Episode %s" % a.week)
        a.mp4 = a.mp4 or find(wk, ["CWR_Ep*_YouTube.mp4", "*.mp4"])
        a.audio = a.audio or find(wk, ["*.wav", "*.WAV", "*.aiff", "*.flac", "*.mp3", "*.m4a"])
    if not a.mp4 or not os.path.exists(a.mp4):
        raise SystemExit("No MP4 to check (looked in %s)" % (wk or "--mp4"))

    print("Checking %s\n" % os.path.relpath(a.mp4, ROOT).replace("\\", "/"))
    v = probe(a.mp4)
    vs, as_ = stream(v, "video"), stream(v, "audio")
    dur = float(v["format"]["duration"])
    fails = []

    def check(label, good, detail):
        print("[%s] %-22s %s" % (OK if good else BAD, label, detail))
        if not good:
            fails.append(label)

    check("resolution", (vs.get("width"), vs.get("height")) == (1920, 1080),
          "%sx%s" % (vs.get("width"), vs.get("height")))
    check("video codec", vs.get("codec_name") == "h264", vs.get("codec_name") or "none")
    check("audio codec", as_.get("codec_name") == "aac", as_.get("codec_name") or "none")

    # Bookends are OVERLAID on the mix, so the runtime must EQUAL the mix. A
    # video longer than its audio means they got concatenated instead.
    if a.audio and os.path.exists(a.audio):
        mix = float(probe(a.audio)["format"]["duration"])
        drift = abs(dur - mix)
        check("duration = the mix", drift <= 2.0,
              "%s vs mix %s (%+.1fs)" % (clock(dur), clock(mix), dur - mix))
    else:
        print("[ ---- ] %-22s %s  (no mix found to compare)" % ("duration", clock(dur)))

    mv = mean_volume(a.mp4)
    check("audio is not silent", mv is not None and mv > -50,
          "mean volume %s dB" % ("?" if mv is None else "%.1f" % mv))

    # Frames to actually look at: a track card, and the last frame — the one
    # place a missing closing line hides.
    shots = os.path.join(os.path.dirname(a.mp4), "_preview")
    os.makedirs(shots, exist_ok=True)
    grabs = [("frame_intro.png", ["-ss", "8"]),
             ("frame_card.png",  ["-ss", str(dur * 0.4)]),
             ("frame_end.png",   ["-sseof", "-1.2"])]
    print()
    for name, seek in grabs:
        dest = os.path.join(shots, name)
        subprocess.run(["ffmpeg", "-v", "error", "-y"] + seek +
                       ["-i", a.mp4, "-frames:v", "1", dest], capture_output=True)
        print("       frame -> %s" % os.path.relpath(dest, ROOT).replace("\\", "/"))

    print()
    if fails:
        print("!! %d check(s) FAILED: %s" % (len(fails), ", ".join(fails)))
        print("   Do not upload this until you know why.")
        return 1
    print("All checks passed. Now OPEN THE THREE FRAMES and look at them —")
    print("ffprobe cannot see a card that drew the wrong text.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
