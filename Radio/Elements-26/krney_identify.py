#!/usr/bin/env python3
"""
krney_identify.py — put artists, genres and release dates on KRNeY's tracklist.

His handwritten sheet has song names and times and NOTHING else: no artists, no
genres. An open web search on a bare title like "Orange" or "Curves" returns
whatever is most popular and would fill the sheet with confident nonsense.

So the search is CONSTRAINED BY THE LINEUP. Every track in an Elements episode
is supposed to be by a producer playing that festival, and we already hold all
157 booked names. A candidate is only accepted when its credited artist is one
of them — which turns an unanswerable question ("who made a song called Orange?")
into an answerable one ("which Elements artist has a song called Orange?").

Anything that doesn't clear that bar is reported as UNIDENTIFIED rather than
guessed. A wrong artist on a card is worse than a blank one, and worse still
here because the whole premise of the episode is that these acts are playing.

    python Radio/Elements-26/krney_identify.py            # dry, prints
    python Radio/Elements-26/krney_identify.py --write    # writes the JSON
"""
import json, os, re, sys, time, urllib.parse, urllib.request

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "render"))
import release_dates as RD

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "render", "krney_identified.json")

# The titles as read off the photo, in play order. Kept in one place so a
# correction to the reading only has to happen here.
TITLES = [
    "Fast lane", "Curves", "Mask off", "Purple Rhythm", "Water Falls", "Evergreen",
    "Orange", "God speed", "Type Shit", "Wet me", "Telekness", "Gimme Gimme",
    "Carry me home", "Feel the vibe", "Forever young", "The Fade", "Feel it in my",
    "Make you Right", "Psycho", "Long walks", "amigud", "gossip", "Hype up",
    "Dream On", "Torch", "Surround sounds", "Let 'em talk",
]
# Artists KRNeY wrote beside two of the rows.
HINTS = {"Torch": "Gorillat", "Let 'em talk": "Mikayli"}


def norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def load_lineup():
    p = os.path.join(HERE, "elements_all_names.json")
    if not os.path.exists(p):
        raise SystemExit("Missing %s — export the lineup names first." % p)
    names = json.load(open(p, encoding="utf-8"))
    idx = {}
    for n in names:
        idx[norm(n)] = n
    return names, idx


# HARD separators always divide two different acts. "&" and "+" are NOT here:
# they more often bind one act's own name ("Above & Beyond", "Chase & Status"),
# and splitting on them is the documented mistake that once dropped Above &
# Beyond's entire catalogue in elements_sc.py.
# NOTE "with" is deliberately absent. It splits band names far more often than it
# separates acts: "Sleeping With Sirens" became ["sleeping", "sirens"], matched the
# booked Sirens, and credited a 2017 hard-rock song to an Elements DJ. The genre
# gate caught that one; the separator is the actual bug.
HARD = re.compile(r"[,/;]| x | vs | versus |\bfeat\.?\b|\bft\.?\b|\bfeaturing\b|\bpres\.?\b|\bpresents\b",
                  re.I)
SOFT = re.compile(r"[&+]")


def credit_segments(credited):
    """The credit cut into acts, hard separators first so duo names survive."""
    parts = [p for p in HARD.split(credited or "") if p.strip()]
    out = [norm(p) for p in parts if norm(p)]
    # only then consider "&" — as a LAST resort, after the intact names failed
    for p in parts:
        out += [norm(x) for x in SOFT.split(p) if norm(x)]
    return out


def lineup_hit(credited, idx):
    """Which booked act is this credit? EXACT segment match only.

    Substring containment was catastrophic here. The lineup includes short names
    — Ade, Hol!, Los, Wan, Freq — and "ade" appears inside Cascade, Shade and
    Adele, "hol" inside Nichols and Holly. That single rule credited "Fast Lane"
    to Ade (a 2011 rock song), "Curves" and "Orange" to Hol!, and "gossip" to
    Sirens as 2017 hard rock. Every one looked like a clean hit in the output.

    So a booked name has to BE one of the acts credited, not appear inside one.
    """
    whole = norm(credited)
    if whole in idx:                       # the credit IS the act, e.g. "Above & Beyond"
        return idx[whole]
    for seg in credit_segments(credited):
        if seg in idx:
            return idx[seg]
    return None


# Genres that an Elements set will not contain. A lineup name can be shared with
# an unrelated band — the booked "Sirens" is a DJ, and a 2017 HARD ROCK song
# called "gossip" by a different Sirens matched perfectly on name and title. The
# festival is electronic; the genre is the last check that catches a namesake.
IMPLAUSIBLE = {"rock", "hard rock", "indie rock", "alternative", "metal", "punk",
               "country", "jazz", "classical", "folk", "blues", "gospel", "opera",
               "films/games", "fusion", "soundtrack", "comedy", "spoken word",
               "children's music", "vocal", "easy listening", "new age", "j-pop",
               "k-pop", "singer/songwriter", "christian & gospel"}


def plausible_genre(g):
    return (g or "").strip().lower() not in IMPLAUSIBLE


def itunes(q):
    js = RD.get("https://itunes.apple.com/search?term=%s&entity=song&limit=25"
                % urllib.parse.quote(q))
    return [{"artist": r.get("artistName", ""), "title": r.get("trackName", ""),
             "date": (r.get("releaseDate") or "")[:10],
             "genre": r.get("primaryGenreName") or "", "src": "iTunes"}
            for r in (js or {}).get("results", [])]


def deezer(q):
    js = RD.get("https://api.deezer.com/search?q=%s&limit=25" % urllib.parse.quote(q))
    out = []
    for r in (js or {}).get("data", []):
        out.append({"artist": (r.get("artist") or {}).get("name", ""),
                    "title": r.get("title", ""), "album_id": (r.get("album") or {}).get("id"),
                    "date": "", "genre": "", "src": "Deezer"})
    return out


def title_matches(want, got):
    """Same song? Titles here are short and generic, so this is deliberately
    strict: exact after normalising, or the read title is the whole first part
    of a longer official title ("Gimme Gimme" vs "Gimme Gimme (Extended Mix)")."""
    w, g = norm(want), norm(RD.clean_title(got))
    if not w or not g:
        return False
    if w == g:
        return True
    return g.startswith(w) and len(w) >= max(6, 0.6 * len(g))


def identify(title, idx, rejected=None):
    rejected = [] if rejected is None else rejected
    hint = HINTS.get(title)
    queries = []
    if hint:
        queries.append("%s %s" % (hint, title))
    queries += [title + " " + "elements", title]
    seen = []
    for q in queries[:1] + queries[1:]:
        for fn in (itunes, deezer):
            try:
                cands = fn(q)
            except Exception:
                cands = []
            for c in cands:
                if not title_matches(title, c["title"]):
                    continue
                who = lineup_hit(c["artist"], idx)
                if not who:
                    continue
                if c["src"] == "Deezer" and c.get("album_id"):
                    a = RD.get("https://api.deezer.com/album/%s" % c["album_id"])
                    c["date"] = ((a or {}).get("release_date") or "")[:10]
                    g = ((a or {}).get("genres") or {}).get("data") or []
                    c["genre"] = g[0].get("name", "") if g else ""
                    time.sleep(0.12)
                c["lineup"] = who
                if not plausible_genre(c.get("genre")):
                    rejected.append((c["artist"], c["title"], c.get("genre"), c["date"]))
                    continue
                seen.append(c)
            time.sleep(0.18)
        if seen:
            break
    if not seen:
        return None
    # earliest credible date wins, same rule as release_dates
    seen.sort(key=lambda c: (c["date"] or "9999"))
    return seen[0]


def main():
    write = "--write" in sys.argv
    names, idx = load_lineup()
    print("matching %d titles against %d booked Elements artists\n" % (len(TITLES), len(names)))
    out, hit = [], 0
    dropped = []
    for i, t in enumerate(TITLES, 1):
        rej = []
        r = identify(t, idx, rej)
        for x in rej:
            dropped.append((t,) + x)
        if r:
            hit += 1
            out.append({"idx": i, "title": t, "artist": r["lineup"], "credited": r["artist"],
                        "release_date": r["date"], "genre": r["genre"], "src": r["src"]})
            print("  %2d  %-18s -> %-22s %-10s %-16s (%s)"
                  % (i, t[:18], r["lineup"][:22], r["date"] or "-", (r["genre"] or "-")[:16], r["src"]))
        else:
            out.append({"idx": i, "title": t, "artist": "", "credited": "",
                        "release_date": "", "genre": "", "src": ""})
            print("  %2d  %-18s -> UNIDENTIFIED" % (i, t[:18]))
        time.sleep(0.2)
    print("\nidentified %d/%d against the lineup" % (hit, len(TITLES)))
    if dropped:
        print("\nrejected on genre — a booked name matched, but not on a record this "
              "festival would play:")
        for t, art, ttl, g, d in dropped:
            print("   %-18s %-20s %-26s %-12s %s" % (t[:18], art[:20], ttl[:26], g, d))
    if write:
        json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("wrote", OUT)
    else:
        print("(dry run - pass --write to save)")


if __name__ == "__main__":
    main()
