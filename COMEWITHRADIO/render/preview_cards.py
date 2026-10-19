#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render an episode's track cards as PNGs, for review before the mix exists.

    python Radio/render/preview_cards.py --txt "Radio/Episode 7/CWR_Ep7_V2.txt" \
        --shows "Radio/Episode 7/EP7_shows.json" --ep "EP 7" \
        --out "Radio/Episode 7/card_previews"

Cards are drawn by render_episode.render_card -- the SAME function the video
uses -- so what is reviewed here is what renders. `--shows` is all_shows.py's
output; any show whose title starts with --spotlight-prefix (default "Dance
Infusion") turns that card into a spotlight card for our own night.
"""
import argparse, json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import render_episode as R
from cues_from_rekordbox import read_rekordbox
from _paths import radio_dir

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


SPOT_PREFIX = "Dance Infusion"
SPOT_LINE = "{artist} plays our fundraiser  ·  tickets at comewith.org"


def card_for(t, s, prefix=SPOT_PREFIX, line=SPOT_LINE):
    """The card fields for one export row + its all_shows entry. Shared with
    cues_with_shows.py, so the video draws exactly what was reviewed here."""
    card = {"artist": t["artist"], "title": t["title"], "genres": t.get("genre", ""),
            "release_date": "", "shows": s["shows"]}
    ours = [x for x in s["shows"] if (x.get("title") or "").startswith(prefix)]
    if ours:
        card["spotlight"] = ours[0]["title"]
        card["shows"] = ours                       # the card leads with OUR night
        card["spotlight_line"] = line.format(artist=(ours[0].get("artists") or [t["artist"]])[0])
    return card


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True)
    ap.add_argument("--shows", required=True, help="all_shows.py JSON")
    ap.add_argument("--out", required=True)
    ap.add_argument("--ep", default="EP 1")
    ap.add_argument("--title", default="Come With NYC Radio")
    ap.add_argument("--cover", default=os.path.join(radio_dir(ROOT), "Artwork", "Radio_Thumbnail.jpg"))
    ap.add_argument("--spotlight-prefix", default=SPOT_PREFIX)
    ap.add_argument("--spotlight-line", default=SPOT_LINE)
    ap.add_argument("--only", help="comma list of track numbers")
    a = ap.parse_args()

    tracks = read_rekordbox(a.txt)
    shows = json.load(open(a.shows, encoding="utf-8"))
    if len(shows) != len(tracks):
        raise SystemExit("shows JSON has %d tracks, export has %d -- re-run all_shows.py"
                         % (len(shows), len(tracks)))
    only = {int(x) for x in a.only.split(",")} if a.only else None
    os.makedirs(a.out, exist_ok=True)
    bg = R.make_background()
    cover = R.rounded_cover(a.cover, 470, 34)
    n = len(tracks)
    for i, (t, s) in enumerate(zip(tracks, shows), 1):
        if only and i not in only:
            continue
        card = card_for(t, s, a.spotlight_prefix, a.spotlight_line)
        out = os.path.join(a.out, "card_%02d%s.png" % (i, "_DI" if card.get("spotlight") else ""))
        R.render_card(bg, cover, card, i, n, a.ep, a.title, i < n, out, progress=i / n)
        print("wrote", out)


if __name__ == "__main__":
    main()
