#!/usr/bin/env python3
"""Sanity-check an episode cover before it gets baked into 27 cards.

    python Radio/render/_check_cover.py "Radio/Episode 3/CWR_EP.3 COVER.JPG"

Never fails the run — the render is still better than no render, and the person
watching the console can decide. It just says out loud what the renderer is
about to do quietly, because both of these have cost a re-render before:

  * NOT SQUARE — render_episode centre-crops to a square, so a portrait cover
    loses its top and bottom on every single card.
  * CMYK — a print colour space. Pillow converts it without an ICC profile, so
    colours shift; Adobe CMYK JPEGs can also come out inverted.
"""
import os, sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

if len(sys.argv) < 2 or not os.path.exists(sys.argv[1]):
    sys.exit(0)

try:
    from PIL import Image
    im = Image.open(sys.argv[1])
except Exception as ex:
    print("               (could not read the cover: %s)" % ex)
    sys.exit(0)

w, h = im.size
notes = []

if w != h:
    lost = abs(w - h)
    pct = 100.0 * lost / max(w, h)
    side = "the sides" if w > h else "the top and bottom"
    if pct >= 5:
        notes.append("NOT SQUARE (%dx%d) — %d px will be cropped off %s of every card."
                     % (w, h, lost, side))
    else:
        notes.append("almost square (%dx%d) — %d px trimmed off %s, not noticeable."
                     % (w, h, lost, side))

if im.mode not in ("RGB", "L"):
    notes.append("colour mode is %s, not RGB — it gets converted, which can shift "
                 "the colours. Check the preview frame." % im.mode)

if min(w, h) < 470:
    notes.append("only %dpx on its short side — it gets scaled UP to 470 and will "
                 "look soft." % min(w, h))

for n in notes:
    print("               ! %s" % n)
