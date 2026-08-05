#!/usr/bin/env python3
"""
krney_align.py — marry KRNeY's handwritten TIMES to the real tracklist he built
in the dashboard, and look up genres + release dates from the result.

Two half-sources:
  * the photo of his notebook has the START TIMES and nothing else usable;
  * the episode's tracks in prod (entered through his DJ link, source='dj') have
    the real ARTISTS and TITLES but no times, and are in his planning order, not
    the order he played.

So they are matched on the song name. That immediately settles most of the nine
readings I could not make out — "Evkvround" is Boys Noize's FVKVRVND, "Telekness"
is Zingara's Telekinesis, "Hype up" is HYYTUP — none of which any amount of
squinting at the page would have produced.

Anything that does not match is printed for a human, both ways round: a written
line with no track, and a track with no written line. Those are the only two
questions left, and they are Keith's or Martin's to answer, not mine to guess.

    python Radio/Elements-26/krney_align.py            # dry
    python Radio/Elements-26/krney_align.py --write    # write the cues CSV
"""
import csv, io, json, os, re, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "render"))
sys.path.insert(0, HERE)
import release_dates as RD
from krney_times_sheet import TRACKS            # the readings live in ONE place

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

OUT = os.path.join(HERE, "render", "Elements_Ep2_KRNeY_cues.csv")


ROOT = os.path.dirname(os.path.dirname(HERE))


def sql_rows(q):
    import urllib.request
    env = {}
    for l in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
        l = l.strip()
        if "=" in l and not l.startswith("#"):
            k, v = l.split("=", 1); env[k] = v.strip().strip('"').strip("'")
    req = urllib.request.Request(
        "https://api.supabase.com/v1/projects/%s/database/query"
        % env.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk"),
        data=json.dumps({"query": q}).encode(),
        headers={"Authorization": "Bearer " + env["SBP_PAT"], "Content-Type": "application/json",
                 "User-Agent": "Mozilla/5.0 Chrome/126"}, method="POST")
    return json.loads(urllib.request.urlopen(req, timeout=60).read().decode() or "[]")


def norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def strip_lead_artist(title, artist):
    """'Chris Lake - Psycho' on Chris Lake's row -> 'Psycho'. Leaves
    'James Hype - Waterfalls (Mersiv Flip)' alone: that names the ORIGINAL."""
    m = re.match(r"^\s*([^()\[\]]{2,50}?)\s+-\s+(.+)$", title or "")
    if m and norm(m.group(1)) == norm(artist):
        return m.group(2).strip()
    return title


def core_title(title):
    """The song's own name, with the packaging taken off.

    "Ganja White Night x Mr. Bill - Mask Off"      -> Mask Off
    "Little Simz - Torch (GorillaT Flip)"          -> Torch
    "ABBA- Gimme! Gimme! Gimme! (Disco Wonk Flip)" -> Gimme! Gimme! Gimme!

    A hand-written line records the SONG, so everything the uploader wrapped
    around it — the original artist in front, the flip credit behind — is noise
    for matching purposes.
    """
    t = re.sub(r"[\(\[][^)\]]*[\)\]]", " ", title or "")
    if " - " in t or "- " in t:
        t = re.split(r"\s*-\s+", t)[-1]
    return t.strip()


def score(read, track_title):
    """How well a handwritten reading matches a real title, 0-100.

    Hand-rolled prefix/containment rules missed the obvious ones: "Surround
    sounds" against "SURROUND SOUND (CRANKDAT REMIX)" failed every branch because
    of a single trailing 's'. A real similarity ratio handles that, and the
    misreadings too — Telekness/Telekinesis, Type Shit/TYPE SH*T.
    """
    import difflib
    a = norm(read)
    for cand in (core_title(track_title), track_title):
        b = norm(cand)
        if not a or not b:
            continue
        if a == b:
            return 100
        r = difflib.SequenceMatcher(None, a, b).ratio()
        if b.startswith(a) or a.startswith(b):
            r = max(r, 0.9)
        if r >= 0.55:
            return int(r * 100)
    return 0


def main():
    write = "--write" in sys.argv
    rows = sql_rows("""
      select t.sort, coalesce(t.artist_name,'') as artist, coalesce(t.title,'') as title
      from sc_playlist_tracks t join sc_playlists p on p.id = t.playlist_id
      where p.edition_name='Come With Elements Radio' and p.edition_seq=2
      order by t.sort;""")
    if not rows:
        raise SystemExit("No tracks on EP 2 in prod.")
    print("%d written times, %d tracks in the dashboard\n" % (len(TRACKS), len(rows)))

    # GLOBAL best-first assignment, not row-by-row greedy. Taking each written
    # line's best free track in order let an early weak match steal a track that
    # was a much stronger match for a later line — "Purple Rhythm" grabbed
    # Psycho, so Zingara's "Purple Plum Trees" was still sitting unclaimed at the
    # end. Scoring every pair and assigning the strongest first fixes that.
    cand = []
    for a, (start, read, sure, note) in enumerate(TRACKS):
        for b, r in enumerate(rows):
            s = max(score(read, strip_lead_artist(r["title"], r["artist"])),
                    score(read, r["title"]))
            if s >= 55:
                cand.append((s, a, b))
    cand.sort(reverse=True)
    take_a, take_b, link = set(), set(), {}
    for s, a, b in cand:
        if a in take_a or b in take_b:
            continue
        take_a.add(a); take_b.add(b); link[a] = (rows[b], s)
    used = take_b
    pairs = [(TRACKS[a][0], TRACKS[a][1]) + (link.get(a, (None, 0)))
             for a in range(len(TRACKS))]

    print("  TIME     WRITTEN              ->  ARTIST                TITLE")
    print("  " + "-" * 88)
    for start, read, r, s in pairs:
        if r:
            print("  %-7s  %-20s ->  %-20s  %s" % (start, read[:20], r["artist"][:20],
                                                   strip_lead_artist(r["title"], r["artist"])[:38]))
        else:
            print("  %-7s  %-20s ->  ** no match **" % (start, read[:20]))
    weak = [(p[0], p[1], p[2]) for p in pairs if p[2] and p[3] < 70]
    if weak:
        print("\n  low-confidence pairings - CONFIRM THESE (the handwriting is the weak link):")
        for start, read, r in weak:
            print("     %-7s  %-20s ->  %s - %s" % (start, read[:20], r["artist"], r["title"][:40]))
    unmatched_tracks = [r for i, r in enumerate(rows) if i not in used]
    print("\nmatched %d/%d written lines" % (sum(1 for p in pairs if p[2]), len(pairs)))
    if unmatched_tracks:
        print("dashboard tracks with no written line (%d):" % len(unmatched_tracks))
        for r in unmatched_tracks:
            print("   %-22s %s" % (r["artist"][:22], r["title"][:50]))

    if not write:
        print("\n(dry run - pass --write to build the cues)")
        return

    # EVERY track goes in the cues, not just the matched ones. Writing only the
    # 22 matched rows and importing that over the station DELETED the four KRNeY
    # entered that no written line reached — Candy Shop Remix, Purple Plum Trees,
    # Hot Mic, FVKVRVND. His list is the record of what he played; the handwriting
    # only supplies times. Unmatched tracks keep their place and go without one.
    out = []
    for start, read, r, s in pairs:
        if not r:
            continue
        out.append({"start": start, "artist": r["artist"],
                    "title": strip_lead_artist(r["title"], r["artist"]),
                    "genres": "", "release_date": "", "conf": s})
    for r in unmatched_tracks:
        out.append({"start": "", "artist": r["artist"],
                    "title": strip_lead_artist(r["title"], r["artist"]),
                    "genres": "", "release_date": "", "conf": 0})
    print("\nlooking up genre + release date for %d tracks (%d carry a time)…"
          % (len(out), len(out) - len(unmatched_tracks)))
    for t in out:
        hits = (RD.from_itunes(t["artist"], t["title"]) or RD.from_deezer(t["artist"], t["title"])
                or RD.from_musicbrainz(t["artist"], t["title"]))
        if hits:
            d, g, src = sorted(hits)[0]
            t["release_date"] = d
            if g:
                t["genres"] = g
            print("   %-20s %-34s %s  %s" % (t["artist"][:20], t["title"][:34], d, g or "-"))
        else:
            print("   %-20s %-34s -" % (t["artist"][:20], t["title"][:34]))
        time.sleep(0.2)

    lineup = json.load(open(os.path.join(HERE, "elements_all_names.json"), encoding="utf-8"))
    lidx = {norm(n): n for n in lineup}
    with io.open(OUT, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot",
                    "genres", "release_date", "show_date", "show_venue", "duration_ms"])
        for i, t in enumerate(out, 1):
            on_bill = norm(t["artist"]) in lidx
            w.writerow([i, t["start"], t["artist"], t["title"], "", "", "",
                        t["genres"], t["release_date"],
                        "2026-08-07" if on_bill else "", "Elements Festival" if on_bill else "",
                        ""])
    print("\nwrote", OUT)
    print("dated %d/%d · genres %d/%d"
          % (sum(1 for t in out if t["release_date"]), len(out),
             sum(1 for t in out if t["genres"]), len(out)))


if __name__ == "__main__":
    main()
