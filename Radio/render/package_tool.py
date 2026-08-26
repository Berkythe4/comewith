#!/usr/bin/env python3
"""
package_tool.py — zip up the render tool so someone else can run it.

    python Radio/render/package_tool.py                     # just the tool
    python Radio/render/package_tool.py --episode 3         # tool + that episode

Produces `CWR_render_tool.zip` you can send to anyone, or carry to another
laptop.

WHAT IT DELIBERATELY DOES NOT INCLUDE
-------------------------------------
`.env`. Never. That file holds SBP_PAT, which is a Supabase *management* token:
it can run arbitrary SQL against production, read every table, and deploy and
delete edge functions. It is the keys to the whole project, not a read-only
database login. Sending it to a guest DJ so they can make a video would be handing
over the entire site.

There is no safer credential to send in its place either — the public key is
blocked from the radio tables by row-level security (verified), so there is no
read-only route that works.

So the package is OFFLINE BY DESIGN. Everything the renderer needs about an
episode is baked into the episode folder by whoever DOES have credentials:

    episode.json      who mixed it, the drop dates, the on-screen label
    EPN_cues.csv      the tracklist, times, genres, and each artist's next show

With those two files present the render needs no network at all. `--episode N`
bundles them for you.

WHAT THEY NEED INSTALLED: Python 3 and ffmpeg. SETUP.txt in the zip says how.
"""
import argparse, io, os, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
from _paths import episode_dir

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

# The tool itself. Listed explicitly rather than globbed: a glob would happily
# sweep up EP*_cues.csv, a stray .env copy, or __pycache__.
TOOL = [
    "render_episode.py", "make_episode.py", "verify_episode.py",
    "tracklist_from_txt.py", "preview_bookends.py", "_paths.py", "make_tags.py",
    "_have_times.py", "_check_cover.py", "templates.json", "tap_times.html",
]
ALSO = [
    ("Radio/Make Radio MP4.bat", "Make Radio MP4.bat"),
    ("Radio/HOW_TO_MAKE_THE_MP4.md", "Radio/HOW_TO_MAKE_THE_MP4.md"),
    ("Radio/Artwork/Radio_Thumbnail.jpg", "Radio/Artwork/Radio_Thumbnail.jpg"),
]
# Files an episode folder needs to render with no database.
EPISODE_KEEP = (".csv", ".json", ".txt", ".png", ".jpg", ".jpeg", ".webp")
EPISODE_SKIP_DIRS = ("_preview",)

SETUP = """COME WITH RADIO — making the episode video
=========================================

You need two things installed. Both are free, both are one command.

1. PYTHON 3      https://www.python.org/downloads/
   On Windows, TICK "Add python.exe to PATH" on the first screen of the
   installer. That checkbox is the whole difference between this working and
   not.

   Then, in a terminal:       pip install pillow

2. FFMPEG        https://ffmpeg.org/download.html
   Windows:      winget install Gyan.FFmpeg
   Mac:          brew install ffmpeg

To check both are ready, open a terminal in this folder and run:

   python --version
   ffmpeg -version

If either says "not recognised", it is not on your PATH — reinstall it with the
PATH option ticked, and open a NEW terminal afterwards.


MAKING A VIDEO
--------------
Double-click `Make Radio MP4.bat` and type the episode number.

Everything it needs is already in the `Radio\\Episode N` folder that came with
this zip. It works completely offline — no logins, no database, no accounts.

The full guide, including how to change any of the wording, is in
Radio\\HOW_TO_MAKE_THE_MP4.md.


CHANGING THE WORDING
--------------------
Radio\\render\\templates.json holds every word the video draws — the opening
slide, the closing slide, and the strip along the bottom of each track card.
Edit it, save it, run the tool again. There is nothing to rebuild.


IF YOU CHANGE THE TRACKLIST
--------------------------
Edit the "Track List ....txt" and run the tool again. It rebuilds the tracklist
from the live site by itself — you do not need an account, a password, or a
login. The episode folder carries a read-only token for that one episode.

What it CANNOT do from here is change the running order on the website. That
needs the real credentials, and stays with whoever runs the site.


WHAT'S NOT IN HERE
------------------
Database credentials, on purpose. Everything you need is already in the episode
folder. If the tool asks for something you don't have, ask whoever sent this —
don't go looking for a login.
"""


def add(z, disk_path, arc_path):
    if os.path.isfile(disk_path):
        z.write(disk_path, arc_path)
        return 1
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", "--week", dest="episode", action="append",
                    help="also bundle this episode's folder (repeatable)")
    ap.add_argument("--out", default=os.path.join(ROOT, "CWR_render_tool.zip"))
    a = ap.parse_args()

    n = 0
    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in TOOL:
            got = add(z, os.path.join(HERE, f), "Radio/render/" + f)
            if not got:
                print("  (missing, skipped: %s)" % f)
            n += got
        for src, arc in ALSO:
            n += add(z, os.path.join(ROOT, src), arc)
        z.writestr("SETUP.txt", SETUP)
        n += 1

        for ep in (a.episode or []):
            folder = episode_dir(ROOT, ep)
            if not folder:
                print("  !! no folder for episode %s — skipped" % ep)
                continue
            meta = os.path.join(folder, "episode.json")
            if not os.path.exists(meta):
                print("  !! episode %s has no episode.json — run make_episode.py "
                      "on a machine WITH credentials first, or it will not render "
                      "on theirs." % ep)
            base = os.path.basename(folder)
            for root, dirs, files in os.walk(folder):
                dirs[:] = [d for d in dirs if d not in EPISODE_SKIP_DIRS]
                for f in files:
                    if not f.lower().endswith(EPISODE_KEEP):
                        continue          # leaves the mix and the mp4 behind
                    p = os.path.join(root, f)
                    rel = os.path.relpath(p, folder).replace("\\", "/")
                    n += add(z, p, "Radio/%s/%s" % (base, rel))
            print("  bundled episode %s (%s) — the mix itself is NOT included, "
                  "send that separately" % (ep, base))

    size = os.path.getsize(a.out) / 1048576.0
    print("\nWrote %s  (%d files, %.1f MB)" % (
        os.path.relpath(a.out, ROOT).replace("\\", "/"), n, size))
    print("No credentials are in this zip. It renders offline.")


if __name__ == "__main__":
    main()
