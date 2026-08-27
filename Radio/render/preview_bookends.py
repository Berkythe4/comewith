#!/usr/bin/env python3
"""
preview_bookends.py — render JUST the intro + a sample song card + the closing,
so the intro/closing design and timing can be reviewed and signed off without
rendering a whole hour-long episode. Reuses the exact functions the real render
uses, so what you approve here is what ships.

    python Radio/render/preview_bookends.py

Writes PNG stills of every reveal beat + a short SILENT mp4 (real timing) into
Radio/Video/_preview/. The mp4 is silent on purpose — this is a look at the
visuals; the finished episode plays them over the mix.
"""
import os, subprocess, sys
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
import render_episode as R

COVER = os.path.join(ROOT, "Radio", "Artwork", "Radio_Thumbnail.jpg")
OUT   = os.path.join(ROOT, "Radio", "Video", "_preview")
os.makedirs(OUT, exist_ok=True)

# Episode meta (EP1, from prod). next_date = the following Thursday's drop.
EP, MIXED_BY, DROP, NEXT = "EP 1", "Berky", "2026-07-23", "2026-07-30"

# A real EP1 track WITH a genre; release_date is filled here so the new
# genre/release line is visible. In production release_date comes from Beatport.
SAMPLE = {"artist": "Sammy Virji", "title": "If U Need It",
          "genres": ["House", "Garage"], "release_date": "2023-11-10",
          "show_date": "2026-07-31", "show_venue": "Brooklyn Army Terminal"}

def main():
    bg = R.make_background()
    cover_sm = R.rounded_cover(COVER, 300, 24)
    cover_lg = R.rounded_cover(COVER, 470, 34)
    frames = []   # (png, dur, label)

    # INTRO beats
    for i, base in enumerate(R.INTRO_BEATS):
        stage = min(i, len(R.INTRO_BEATS) - 2)
        p = os.path.join(OUT, "intro_%02d.png" % i)
        R.draw_intro(bg, cover_sm, EP, MIXED_BY, DROP, stage).save(p)
        frames.append((p, base, "intro-beat-%d" % i))

    # SAMPLE song card (shows the new genre + release-date line)
    p = os.path.join(OUT, "song_card.png")
    R.render_card(bg, cover_lg, SAMPLE, 8, 18, EP, "Come With Radio",
                  "Luke Dean — This Heat", p, progress=0.42)
    frames.append((p, 4.0, "song-card"))

    # CLOSING beats
    # R.outro_beats(), never a local copy — the preview must show exactly what
    # the render will draw, including a dropped beat when the tease is blank.
    _beats = R.outro_beats()
    for i, base in enumerate(_beats):
        stage = min(i, len(_beats) - 2)
        p = os.path.join(OUT, "outro_%02d.png" % i)
        R.draw_outro(bg, cover_sm, NEXT, stage).save(p)
        frames.append((p, base, "closing-beat-%d" % i))

    # silent mp4 at real timing
    listfile = os.path.join(OUT, "list.txt")
    with open(listfile, "w", encoding="utf-8") as f:
        for png, dur, _ in frames:
            f.write("file '%s'\n" % png.replace("\\", "/"))
            f.write("duration %.3f\n" % dur)
        f.write("file '%s'\n" % frames[-1][0].replace("\\", "/"))
    mp4 = os.path.join(OUT, "CWR_intro_closing_preview.mp4")
    cmd = ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", listfile,
           "-vf", "fps=30,format=yuv420p", "-c:v", "libx264", "-preset", "veryfast",
           "-crf", "20", "-pix_fmt", "yuv420p", "-movflags", "+faststart", mp4]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        sys.stderr.write(r.stderr[-1200:]); raise SystemExit("ffmpeg failed")

    print("Wrote %d stills + preview mp4 to %s" % (len(frames), os.path.relpath(OUT, ROOT)))
    print("MP4:", os.path.relpath(mp4, ROOT).replace("\\", "/"))

if __name__ == "__main__":
    main()
