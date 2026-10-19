#!/usr/bin/env python3
"""
cues_from_soundcloud.py — build the render cues + the start-times fill-in sheet
from a SoundCloud PLAYLIST (a "set").

    python Radio/render/cues_from_soundcloud.py \
        --url "https://soundcloud.com/henryspostingmusic/sets/elements-saturday" \
        --out   "Radio/Elements-26/render/Elements_Ep3_cues.csv" \
        --times "Radio/Elements-26/render/Elements_Ep3_times.txt" \
        --ep 3 --show "COME WITH ELEMENTS RADIO"

Same output as cues_from_rekordbox.py — it imports that module's sheet writer
rather than growing a second, drifting copy of it. Use this when the set was
assembled on SoundCloud instead of exported from Rekordbox.

Three things this has to get right:

  * TRACKING PARAMS. A share link arrives as "?si=…&utm_source=clipboard&…".
    /resolve wants the bare permalink; the junk also makes the URL unusable as a
    stored value. Stripped before anything else.

  * STUBS. A playlist payload hydrates only the first handful of tracks — this
    set returned 5 of 26 with titles and 21 as bare {id, kind}. A reader that
    trusts the payload silently loses 80% of the set. /tracks?ids= resolves them
    50 at a time, in the playlist's own order.

  * WHO THE ARTIST IS. The uploader is not the artist: "Prospa & Cloonee - Free
    Your Mind" sits on the CircoLoco Records account, and its rights credit says
    CircoLoco Records too. So an "Artist - Title" title is split first, then the
    rights credit, then the uploader — see credited().
"""
import argparse, csv, io, json, os, re, sys, time, urllib.parse, urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cues_from_rekordbox import mmss, write_times_sheet, make_show_lookup

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

API = "https://api-v2.soundcloud.com"
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}
ROOT = r"C:\Users\Admin\Documents\Comewith"


def client_id():
    """The public client_id the site already uses, from site_content."""
    env = {}
    for l in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
        l = l.strip()
        if "=" in l and not l.startswith("#"):
            k, v = l.split("=", 1); env[k] = v.strip().strip('"').strip("'")
    req = urllib.request.Request(
        "https://api.supabase.com/v1/projects/%s/database/query"
        % env.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk"),
        data=json.dumps({"query": "select value from site_content where key='ops.sc_client_id';"}).encode(),
        headers={"Authorization": "Bearer " + env["SBP_PAT"],
                 "Content-Type": "application/json", **UA}, method="POST")
    return json.loads(urllib.request.urlopen(req, timeout=40).read().decode())[0]["value"]


def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
        return json.load(r)


def clean_url(u):
    """Drop the share tracking (?si=, utm_*) — /resolve wants the bare permalink."""
    p = urllib.parse.urlsplit(u.strip())
    return urllib.parse.urlunsplit((p.scheme or "https", p.netloc, p.path, "", ""))


def norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def account_name(username):
    """A SoundCloud display name, minus the account decoration.

    Artist names are reproduced exactly — Theø stays Theø. What comes off is the
    dressing people hang on a HANDLE: a trailing emoji ("RAY VOLPE 🤖") or a
    bracketed tagline, same call already made for "KETTAMA (G-TOWN FOREVER)".
    Only ever trimmed from the ends, never from inside the name.
    """
    s = (username or "").strip()
    s = re.sub(r"\s*[\(\[][^)\]]{0,40}[\)\]]\s*$", "", s)          # trailing tagline
    s = re.sub(r"^[\s\W_]*(?=[A-Za-z0-9])", "", s)                  # leading decoration
    s = re.sub(r"[\s‍️]*[^\w\sÀ-ɏ&.'’!?+-]+\s*$", "", s)  # trailing emoji/symbols
    return s.strip() or (username or "").strip()


def credited(t):
    """(artist, title) — the record's own credit, not the uploader's account.

    An "Artist - Title" title WINS over publisher_metadata. That looks backwards
    next to the ownership rule elsewhere, and it isn't: the rights field names
    the rights holder, which on a label upload is the label. CircoLoco Records
    had "Prospa & Cloonee - Free Your Mind" credited to CircoLoco Records — true
    as rights, wrong on a card that is supposed to name who is playing.
    """
    title = (t.get("title") or "").strip()
    pm = ((t.get("publisher_metadata") or {}).get("artist") or "").strip()
    uploader = account_name((t.get("user") or {}).get("username") or "")
    flat = norm(title)
    up = norm(uploader)
    # The uploader's own name inside the title means THEY made this version:
    # "DAFT PUNK - AROUND THE WORLD (WESTEND EDIT)" on Westend's account is a
    # Westend record. Crediting Daft Punk there would put an Elements chip on an
    # act who isn't playing it — Westend is.
    if up and len(up) >= 4 and up in flat:
        m = re.match(r"^\s*([^()\[\]]{2,60}?)\s+-\s+(.+)$", title)
        # "Subtronics - Fibonacci (Part 1)" on Subtronics' account: take the
        # prefix as the credit, or the sheet reads "Subtronics - Subtronics -
        # Fibonacci". STARTS-WITH, not equals, so a collaboration keeps everyone:
        # "Subtronics x A Little Sound", "Hedex, Sullivan King, Basslayerz".
        # "DAFT PUNK - AROUND THE WORLD (WESTEND EDIT)" keeps its prefix in the
        # title — that names the original, which is not the act playing.
        if m and norm(m.group(1)).startswith(up):
            return m.group(1).strip(), m.group(2).strip()
        return uploader, title
    m = re.match(r"^\s*([^()\[\]]{2,60}?)\s+-\s+(.+)$", title)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return (pm or uploader), title


def fetch(url, cid):
    pl = get("%s/resolve?url=%s&client_id=%s" % (API, urllib.parse.quote(clean_url(url), safe=""), cid))
    if pl.get("kind") != "playlist":
        raise SystemExit("That link is a %s, not a playlist." % pl.get("kind"))
    items = pl.get("tracks") or []
    order = [str(t.get("id")) for t in items if t.get("id")]
    have = {str(t["id"]): t for t in items if t.get("id") and t.get("title")}
    missing = [i for i in order if i not in have]
    for k in range(0, len(missing), 50):
        batch = missing[k:k + 50]
        for t in (get("%s/tracks?ids=%s&client_id=%s" % (API, ",".join(batch), cid)) or []):
            if t.get("id"):
                have[str(t["id"])] = t
        time.sleep(0.15)
    if missing:
        print("hydrated %d stub track(s) the playlist payload left empty" % len(missing))
    out = []
    for i in order:
        t = have.get(i)
        if not t:
            print("  !! could not resolve track id %s — it may have been deleted" % i)
            continue
        artist, title = credited(t)
        out.append({"artist": artist, "title": title,
                    "genre": (t.get("genre") or "").strip(),
                    "bpm": "", "camelot": "",
                    "duration_ms": int(t.get("duration") or 0),
                    "permalink_url": t.get("permalink_url") or ""})
    return pl, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="the SoundCloud playlist (set) link")
    ap.add_argument("--out", required=True, help="cues CSV to write")
    ap.add_argument("--times", help="also write the human fill-in sheet here")
    ap.add_argument("--ep", default="?", help="episode number for the sheet header")
    ap.add_argument("--show", default="COME WITH NYC RADIO", help="show name in the sheet header")
    ap.add_argument("--mix-mins", type=float,
                    help="length of the RECORDED mix in minutes, if you already know it — "
                         "the guesses are scaled to it. Default: the sum of the track lengths.")
    ap.add_argument("--show-venue", default="")
    ap.add_argument("--show-dates", help="JSON {artist: 'YYYY-MM-DD'} for the date chip")
    a = ap.parse_args()

    pl, tr = fetch(a.url, client_id())
    if not tr:
        raise SystemExit("No tracks resolved from that playlist.")
    print("%s — %s · %d tracks" % (pl.get("title"), (pl.get("user") or {}).get("username"), len(tr)))

    total = sum(t["duration_ms"] for t in tr) / 1000.0
    mix = a.mix_mins * 60 if a.mix_mins else total
    scale = (mix / total) if total else 1.0
    run, starts = 0.0, []
    for t in tr:
        starts.append(mmss(run))
        run += t["duration_ms"] / 1000.0 * scale

    show_date_for = make_show_lookup(a.show_dates)

    matched = 0
    with io.open(a.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "genres", "release_date", "show_date", "show_venue", "duration_ms"])
        for i, t in enumerate(tr, 1):
            sd = show_date_for(t["artist"])
            if sd:
                matched += 1
            w.writerow([i, starts[i - 1], t["artist"], t["title"], "", "", "",
                        t["genre"], "", sd, a.show_venue if (sd or a.show_venue) else "",
                        t["duration_ms"]])
    print("wrote %s" % a.out)
    print("track lengths total %s%s" % (mmss(total),
          (" | scaled to a %s mix" % mmss(mix)) if a.mix_mins else " (no mix length given — guesses use this)"))
    if a.show_dates:
        print("show dates matched: %d/%d" % (matched, len(tr)))
        for t in tr:
            if not show_date_for(t["artist"]):
                print("   no show date: %s" % t["artist"])

    if a.times:
        bad = write_times_sheet(
            a.times, tr, starts, a.ep, mix,
            "%s (SoundCloud set - %d tracks, play order)" % (pl.get("title") or "playlist", len(tr)),
            "not recorded yet" if not a.mix_mins else "%s min" % a.mix_mins,
            show=a.show)
        print("wrote %s (non-ASCII chars: %d)" % (a.times, bad))


if __name__ == "__main__":
    main()
