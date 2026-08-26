#!/usr/bin/env python3
"""
make_tags.py — the tag list to paste into YouTube and SoundCloud.

    python Radio/render/make_tags.py --episode 3

Writes `Radio/Episode N/EPN_tags.txt`: the static tags that go out on every
episode, then the genres actually played in THIS one, most common first.

Reads the cues CSV, so it needs no database and works from a folder that was
sent to someone — the same file the renderer draws the cards from.

THE TWO PLATFORMS DISAGREE, so the file gives you a separate ready-to-paste
line for each rather than one list that is subtly wrong on both:

  YouTube      comma separated, and the whole box is capped at 500 CHARACTERS.
               Over that, YouTube silently drops the overflow — you paste 40
               tags, it keeps 28, and nothing tells you.
  SoundCloud   space separated, and a multi-word tag must be "in quotes" or it
               becomes several one-word tags. Capped at 30 tags.

Anything that will not fit is reported rather than quietly dropped.
"""
import argparse, csv, io, os, re, sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, HERE)
from _paths import episode_dir

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

YT_CHAR_CAP = 500
SC_TAG_CAP = 30


def load_static():
    """Static tags live in templates.json so they are editable without code."""
    import json
    try:
        with io.open(os.path.join(HERE, "templates.json"), encoding="utf-8") as f:
            return [t for t in (json.load(f).get("tags") or {}).get("static", []) if t.strip()]
    except Exception as ex:
        print("(templates.json unreadable — no static tags: %s)" % ex)
        return []


def split_genres(raw):
    if isinstance(raw, list):
        parts = raw
    else:
        parts = re.split(r"[|,/·]", str(raw or ""))
    return [p.strip() for p in parts if p and p.strip()]


def titlecase(g):
    """'tech house' -> 'Tech House', but leave 'R&B' and 'UK Garage' alone."""
    if g.isupper() or any(c.isdigit() for c in g):
        return g
    return " ".join(w if (w.isupper() and len(w) > 1) else w.capitalize()
                    for w in g.split())


def dedupe(tags):
    """Case-insensitive, order-preserving."""
    seen, out = set(), []
    for t in tags:
        k = re.sub(r"[^a-z0-9]+", "", t.lower())
        if k and k not in seen:
            seen.add(k)
            out.append(t)
    return out


def yt_line(tags):
    """Pack tags into YouTube's 500-character budget, in order."""
    kept, used = [], 0
    for t in tags:
        add = len(t) + (2 if kept else 0)          # ", "
        if used + add > YT_CHAR_CAP:
            continue                                # try the next, shorter one
        kept.append(t)
        used += add
    return ", ".join(kept), kept


def sc_line(tags):
    kept = tags[:SC_TAG_CAP]
    return " ".join(('"%s"' % t) if " " in t else t for t in kept), kept


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--episode", "--week", dest="episode", required=True)
    ap.add_argument("--cues", help="cues CSV (default: the one in the episode folder)")
    ap.add_argument("--out", help="where to write (default: EPN_tags.txt in the folder)")
    a = ap.parse_args()

    folder = episode_dir(ROOT, a.episode)
    if not folder:
        raise SystemExit("No folder for episode %s" % a.episode)

    cues = a.cues
    if not cues:
        import glob
        hits = sorted(glob.glob(os.path.join(folder, "EP*_cues.csv")))
        if not hits:
            raise SystemExit("No cues CSV in %s — build the tracklist first." % folder)
        cues = hits[0]

    rows = list(csv.DictReader(io.open(cues, encoding="utf-8-sig")))
    counts = Counter()
    for r in rows:
        for g in split_genres(r.get("genres")):
            counts[titlecase(g)] += 1
    genres = [g for g, _ in counts.most_common()]
    artists = dedupe([(r.get("artist") or "").strip() for r in rows if (r.get("artist") or "").strip()])

    static = load_static()
    ordered = dedupe(static + genres)               # static ALWAYS first
    yt, yt_kept = yt_line(ordered)
    sc, sc_kept = sc_line(ordered)

    dropped_yt = [t for t in ordered if t not in yt_kept]
    dropped_sc = [t for t in ordered if t not in sc_kept]

    out = a.out or os.path.join(folder, "EP%s_tags.txt" % a.episode)
    L = []
    L.append("TAGS — Episode %s" % a.episode)
    L.append("=" * 46)
    L.append("")
    L.append("Genres played: %s" % (", ".join("%s (%d)" % (g, counts[g]) for g in genres) or "none in the cues"))
    L.append("")
    L.append("-- YOUTUBE ------------------------------------")
    L.append("Paste into the Tags box. %d tags, %d/%d characters." % (len(yt_kept), len(yt), YT_CHAR_CAP))
    L.append("")
    L.append(yt)
    L.append("")
    if dropped_yt:
        L.append("Did not fit in YouTube's 500 characters: " + ", ".join(dropped_yt))
        L.append("")
    L.append("-- SOUNDCLOUD ---------------------------------")
    L.append("Paste into Additional tags. %d of a maximum %d tags." % (len(sc_kept), SC_TAG_CAP))
    L.append('Multi-word tags are quoted so SoundCloud keeps them whole.')
    L.append("")
    L.append(sc)
    L.append("")
    if dropped_sc:
        L.append("Over SoundCloud's 30-tag limit, so left out: " + ", ".join(dropped_sc))
        L.append("")
    L.append("-- OPTIONAL: the artists ----------------------")
    L.append("Not included above. Good for discovery, but they are names other")
    L.append("people search for, so use them only where that is fair game.")
    L.append("")
    L.append(", ".join(artists))
    L.append("")
    io.open(out, "w", encoding="utf-8").write("\n".join(L))

    print("Wrote %s" % os.path.relpath(out, ROOT).replace("\\", "/"))
    print("  %d static + %d genre tags" % (len(static), len(genres)))
    print("  YouTube    %d tags, %d/%d chars%s" % (
        len(yt_kept), len(yt), YT_CHAR_CAP, "" if not dropped_yt else "  (%d did not fit)" % len(dropped_yt)))
    print("  SoundCloud %d/%d tags%s" % (
        len(sc_kept), SC_TAG_CAP, "" if not dropped_sc else "  (%d over the limit)" % len(dropped_sc)))


if __name__ == "__main__":
    main()
