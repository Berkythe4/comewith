#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Put a station's tracklist into the order that was actually PLAYED.

The dashboard's "Import Rekordbox order" is the normal way to do this. This is
the same operation for when the dashboard isn't the right tool -- a set arranged
outside the SoundCloud playlist entirely, so there's nothing to sync back, and
the export has to be reconciled against the station from here.

    python Radio/render/apply_station_from_rekordbox.py \
        --txt "Radio/Episode 4/CWR_Ep4_rekordbox.txt" --station 8          # review
    python Radio/render/apply_station_from_rekordbox.py \
        --txt "Radio/Episode 4/CWR_Ep4_rekordbox.txt" --station 8 --apply  # write

Prints the full reorder / insert / drop plan and exits. Nothing is written
without --apply, and --apply re-prints the plan and refuses on any disagreement
it can detect. Every statement goes through db.py, so the project ref is visible
in the command being approved.

WHAT IT DOES, and why each part is the way it is:

  * MATCHED rows keep their database row -- and with it the genres, listen link
    and buy link that were researched into it. Only sort, bpm, key and (unlike
    the dashboard importer) title/artist are rewritten, because Rekordbox owns
    what was played and the SoundCloud pull's titles are handles with the artist
    folded in ("Flyinglotus", "leeburridge - LOST IN A MOMENT - MATTHEW DEKAY
    and LEE BURRIDGE - INNERVISIONS").

  * The show chip is RE-DERIVED from ra_artists for every played row, kept ones
    included, and cleared when nobody credited has an upcoming NYC date. The
    chip claims "this artist is playing NYC on this date", so it goes stale on
    its own as dates pass -- SHOW 8 was carrying a 2026-09-06 booking that had
    already happened. A card with no show line is the honest answer when there
    is no show; a date nobody is playing is not. LEARNINGS SS26.

  * DROPPED rows are logged to sc_song_log as passed@N first, which is what
    makes them carry into the next station at finalize. Delete without that and
    the song is simply forgotten. Same call the dashboard's raLogPassed makes.

  * Matching carries the guards CLAUDE.md records for store matching, because
    the traps are identical: "(Original Mix)" / "(Extended Mix)" are QUALIFIERS
    and must still match their own release; a named remixer has to match on both
    sides or an original matches its remix; and a match may never rest on the
    artist alone (a bare artist name inside a long title is not a title match).
"""
import argparse, io, json, os, re, subprocess, sys, unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rekordbox_clean import clean_artist, clean_title

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PROD = "yaytdosxfhcqatmhctzk"


# ---- db.py wrapper -----------------------------------------------------------
def sql(q, ref=PROD):
    """One db.py call. The ref is passed as a LITERAL, per CLAUDE.md -- the whole
    point is that the target project is visible in the command being approved.

    Anything long goes via a temp .sql FILE, which db.py accepts as its argument.
    Windows caps a command line at ~32k, and a full station rewrite is ~100
    statements: passed inline it dies with WinError 206 before the process even
    starts. That failure is safe (nothing ran) but it would cap this tool at
    about 40 tracks, silently, for no reason.
    """
    env = dict(os.environ, SBP_REF=ref)
    tmp = None
    try:
        if len(q) > 8000:
            tmp = os.path.join(REPO, ".apply_station.tmp.sql")
            io.open(tmp, "w", encoding="utf-8").write(q)
            arg = tmp
        else:
            arg = q
        p = subprocess.run([sys.executable, os.path.join(REPO, "db.py"), arg],
                           capture_output=True, text=True, env=env, cwd=REPO, encoding="utf-8")
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    if p.returncode != 0:
        raise SystemExit("db.py failed:\n%s" % (p.stderr or p.stdout))
    return json.loads(p.stdout)


def lit(v):
    if v is None or v == "":
        return "null"
    return "'" + str(v).replace("'", "''") + "'"


# ---- parsing -----------------------------------------------------------------
def read_rekordbox(path):
    """UTF-16 + BOM, tab separated, columns BY HEADER NAME. Same traps the
    dashboard importer and cues_from_rekordbox.py already learned."""
    raw = open(path, "rb").read()
    for enc in ("utf-16", "utf-16-le", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            if "\t" in text:
                break
        except Exception:
            continue
    else:
        raise SystemExit("Could not decode %s" % path)
    lines = [l for l in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if l.strip()]
    head = [h.strip().lstrip("﻿").strip('"').lower() for h in lines[0].split("\t")]
    idx = {h: i for i, h in enumerate(head)}

    def col(cells, *names):
        for n in names:
            i = idx.get(n)
            if i is not None and i < len(cells) and cells[i].strip():
                return cells[i].strip()
        return ""

    out = []
    for ln in lines[1:]:
        c = ln.split("\t")
        title = col(c, "track title", "title", "name")
        if not title:
            continue
        artist = col(c, "artist", "artists")
        if not artist and " - " in title:            # artist folded into the title
            artist, title = [p.strip() for p in title.split(" - ", 1)]
        mix = col(c, "mix name", "mix")
        # Rekordbox keeps "Mix Name" in its own column; fold it back unless the
        # title already says it, or it prints the qualifier twice.
        if mix and mix.lower() not in title.lower() and not re.match(r"^original(\s*mix)?$", mix, re.I):
            title += " (%s)" % mix
        title, artist = clean_title(title), clean_artist(artist)
        t = col(c, "time", "duration", "length")
        secs = 0.0
        if ":" in t:
            p = [float(x) for x in t.split(":")]
            secs = p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]
        bpm = col(c, "bpm")
        try: bpm = int(round(float(bpm))) if bpm else None
        except Exception: bpm = None
        out.append({"title": title.replace("﻿", ""), "artist": artist.replace("﻿", ""),
                    "bpm": bpm, "key": col(c, "key", "tonality"), "genre": col(c, "genre"),
                    "label": col(c, "label"), "duration_ms": int(secs * 1000)})
    return out


# ---- matching ----------------------------------------------------------------
DIAC = re.compile(r"[̀-ͯ]")
NOISE = re.compile(r"\b(original mix|extended mix|extended|radio edit|feat|ft|featuring|the)\b")
QUALIFIER = re.compile(r"^(original|extended|extended revisit|revisit|radio|club|instrumental|"
                       r"dub|vocal|video|album|single|edit)$", re.I)


def norm(s):
    s = DIAC.sub("", unicodedata.normalize("NFD", str(s or "").lower()))
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", NOISE.sub(" ", s)).strip()


def remixer(s):
    """Who a remix is BY. Empty for a plain qualifier -- "(Extended Mix)" is not
    a remix, and treating it as one stops a track matching itself."""
    out = set()
    for x in re.findall(r"[\(\[]([^)\]]*?)\s*(?:remix|rework|edit|bootleg|flip|blend|mix|version)"
                        r"[^)\]]*[\)\]]", s or "", re.I):
        x = re.sub(r"\b(feat|ft|featuring)\b.*$", "", x, flags=re.I).strip()
        if x and not QUALIFIER.match(x) and norm(x):
            out.add(norm(x))
    return out


def toks(s):
    return set(norm(s).split())


def score(r, d):
    a, b = toks(r["title"]), toks(d["title"] or "")
    if not a or not b:
        return 0.0
    b = b | toks(d["artist_name"] or "")          # DB titles often carry the artist
    ov = len(a & b) / len(a)
    ra, rd = remixer(r["title"]), remixer(d["title"] or "")
    if (ra or rd) and not (ra & rd):
        ov *= 0.35                                 # a remix and its original differ
    aa = toks(r["artist"])
    ab = toks(d["artist_name"] or "") | toks(d["title"] or "")
    sq = norm(d["artist_name"] or "").replace(" ", "") + " " + norm(d["title"] or "").replace(" ", "")
    if aa and not (aa & ab) and not any(w in sq for w in aa if len(w) > 3):
        ov *= 0.5                                  # the artist has to corroborate
    return ov


CAMELOT = re.compile(r"^(\d{1,2})\s*([ABab])$")
PITCH = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
CAM_MAJ = {0: "8B", 1: "3B", 2: "10B", 3: "5B", 4: "12B", 5: "7B", 6: "2B", 7: "9B", 8: "4B", 9: "11B", 10: "6B", 11: "1B"}
CAM_MIN = {0: "5A", 1: "12A", 2: "7A", 3: "2A", 4: "9A", 5: "4A", 6: "11A", 7: "6A", 8: "1A", 9: "8A", 10: "3A", 11: "10A"}
CAM_TO_KEY = {}
for p, c in CAM_MAJ.items(): CAM_TO_KEY[c] = PITCH[p]
for p, c in CAM_MIN.items(): CAM_TO_KEY[c] = PITCH[p] + "m"


def key_fields(k):
    """Rekordbox shows key musically ("F#m") OR as Camelot ("8A") depending on
    preference. Accept both, fill the other side in -- same as raKeyFields."""
    k = str(k or "").strip()
    if not k:
        return None, None
    m = CAMELOT.match(k)
    if m:
        c = m.group(1) + m.group(2).upper()
        return CAM_TO_KEY.get(c), c
    kk = k.replace("maj", "").replace("Maj", "").strip()
    for p, name in enumerate(PITCH):
        for suffix, table in ((("m", "min"), CAM_MIN), (("",), CAM_MAJ)):
            for sfx in suffix:
                if kk.lower() == (name + sfx).lower():
                    return kk, table[p]
    return k, None


SPLIT = re.compile(r"\s*(?:,|&|/|\+|\bx\b|\bvs\.?\b|\bb2b\b|\bversus\b|\bfeat\.?\b|\bft\.?\b|"
                   r"\bwith\b|\band\b|\bpres\.?\b)\s*", re.I)


def credits_of(row):
    """Every name this track might be filed under: the credit field split into
    individual artists, plus whoever the remix is by. A remixer counts -- the
    chip answers "why is this song in the show", and NOTES SS4 already says to
    use the other partner's booking when one half has no NYC date."""
    names = [p.strip() for p in SPLIT.split(row["artist"] or "") if len(p.strip()) > 1]
    for x in re.findall(r"[\(\[]([^)\]]*?)\s*(?:remix|rework|edit|mix|version)[^)\]]*[\)\]]",
                        row["title"] or "", re.I):
        x = re.sub(r"\b(extended|original|radio|club|feat|ft|featuring)\b", " ", x, flags=re.I).strip()
        if len(x) <= 1 or QUALIFIER.match(x):
            continue
        # "(Sonny Fodera Extended Deep Mix)" leaves "Sonny Fodera Deep" -- the
        # name plus a genre word the strip list can't know about. Offer the
        # leading prefixes too, longest first. They are only ever matched against
        # real ra_artists rows, so a prefix that names nobody simply misses.
        w = x.split()
        names.extend(" ".join(w[:k]) for k in range(len(w), 0, -1))
    seen, out = set(), []
    for n in names:
        k = norm(n).replace(" ", "")
        if k and k not in seen:
            seen.add(k); out.append(n)
    return out


def show_chips(rows, ref):
    """artist-key -> (date, venue) for every credited name with a FUTURE NYC show.

    Read from ra_events.lineup, NOT from ra_artists.next_event_date.

    `next_event_date` is ONE denormalised field standing in for a plural
    question, and it goes stale the moment a show passes without a re-pull: on
    2026-09-09, 183 artists carried a next_event_date that was wrong (169 already
    in the past, 6 null, 8 later than their real next show) while genuinely
    having an upcoming NYC date. Kim Anh is the worked example -- her row said
    2026-09-06 Gabriela, which had happened, while ra_events had her at Nowadays
    on the 12th, BASEMENT on the 19th, Knockdown on the 25th and Paragon on
    2026-10-09. Reading the singular field, this tool cleared her chip and
    reported "no upcoming NYC date for anyone credited", which was false.

    Same shape as the `next_venue` trap in CLAUDE.md: the tell is a singular
    field backing a plural question. The lineup index is the real answer.

    A track whose artists genuinely have no upcoming show still gets NO chip --
    render_card omits that line, which is the honest result.
    """
    want = set()
    for r in rows:
        for n in credits_of(r):
            want.add(norm(n).replace(" ", ""))
    if not want:
        return {}
    vals = ", ".join("(%s)" % lit(w) for w in sorted(want))
    # Prefer a named room over "TBA" on the same date -- both are true, one is
    # useful on a card. Venue comes from the resolved venues row where 208 gave
    # us one, so the chip prints the canonical spelling of the room.
    q = ("with want(k) as (values %s), "
         "ev as ("
         "  select e.event_date,"
         "         coalesce(v.name, e.venue_name) as venue,"
         "         lower(regexp_replace(l->>'name','[^a-zA-Z0-9]','','g')) as k"
         "    from ra_events e"
         "    left join venues v on v.id = e.venue_id"
         "    cross join lateral jsonb_array_elements(coalesce(e.lineup,'[]'::jsonb)) l"
         "   where e.event_date >= current_date"
         "     and coalesce(v.name, e.venue_name) is not null) "
         "select distinct on (w.k) w.k, ev.event_date, ev.venue "
         "  from want w join ev on ev.k = w.k "
         " order by w.k, ev.event_date, (ev.venue ilike 'TBA') , ev.venue;" % vals)
    return {r["k"]: (r["event_date"], r["venue"]) for r in sql(q, ref)}


def chip_for(row, chips):
    for n in credits_of(row):
        hit = chips.get(norm(n).replace(" ", ""))
        if hit:
            return hit
    return (None, None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True, help="Rekordbox playlist export")
    ap.add_argument("--station", type=int, required=True, help="station_no (the SHOW number)")
    ap.add_argument("--apply", action="store_true", help="actually write to prod")
    ap.add_argument("--keep-extras", action="store_true",
                    help="don't drop the songs that weren't played; trail them after")
    ap.add_argument("--ref", default=PROD)
    a = ap.parse_args()

    rb = read_rekordbox(a.txt)
    if not rb:
        raise SystemExit("No tracks parsed -- check the export.")

    pls = sql("select id, station_no, name, status, published from sc_playlists "
              "where station_no = %d;" % a.station, a.ref)
    if not pls:
        raise SystemExit("No station with station_no=%d" % a.station)
    pl = pls[0]
    if pl["status"] in ("live", "archived"):
        raise SystemExit("SHOW %s is %s -- reopen it (status back to testing) before reordering."
                         % (pl["station_no"], pl["status"]))

    db = sql("select id, sort, sc_track_id, artist_name, title, source, duration_ms, "
             "permalink_url, artwork_url, buy_url, label, show_date, show_venue "
             "from sc_playlist_tracks where playlist_id = '%s' order by sort;" % pl["id"], a.ref)

    plan = {"reorder": [], "insert": [], "drop": []}
    used = set()
    for i, r in enumerate(rb, 1):
        r["n"] = i
        best, bs = None, 0.0
        for d in db:
            if d["id"] in used:
                continue
            s = score(r, d)
            if s > bs:
                best, bs = d, s
        if best and bs >= 0.62:
            used.add(best["id"])
            plan["reorder"].append({"row": r, "db": best, "score": round(bs, 2)})
        else:
            plan["insert"].append({"row": r, "score": round(bs, 2)})
    plan["drop"] = [d for d in db if d["id"] not in used]

    # Re-derive the show chip for every played row from ra_artists. Doing this
    # for the KEPT rows too is the point: a chip written weeks ago can have gone
    # stale, and a date that has already passed on a "catch them live" card is
    # worse than no card line at all.
    chips = show_chips(rb, a.ref)
    for m in plan["reorder"] + plan["insert"]:
        m["chip"] = chip_for(m["row"], chips)

    # ---- review -------------------------------------------------------------
    print("=" * 96)
    print("SHOW %s -- %s  (%s)" % (pl["station_no"], pl["name"], pl["status"]))
    print("%d rows in the station, %d in the export" % (len(db), len(rb)))
    print("=" * 96)
    def chipstr(m, old=None):
        d, v = m["chip"]
        if d:
            was = ""
            if old and old[0] and old[0] != d:
                was = "  (was %s %s)" % (old[0], old[1] or "")
            return "%s %s%s" % (d, v, was)
        if old and old[0]:
            return "CLEARED -- %s %s is in the past and nothing replaces it" % (old[0], old[1] or "")
        return "-- no upcoming NYC date for anyone credited"

    print("\nKEPT + reordered  (%d)   -- row kept, so genres/links survive\n" % len(plan["reorder"]))
    for m in plan["reorder"]:
        r, d = m["row"], m["db"]
        print("  %2d  %.62s" % (r["n"], r["artist"] + " - " + r["title"]))
        print("      was  %-58.58s  %.2f" % ((d["artist_name"] or "") + " - " + (d["title"] or ""), m["score"]))
        print("      show %s" % chipstr(m, (d["show_date"], d["show_venue"])))
    print("\nNEW  (%d)   -- inserted as source='rekordbox'\n" % len(plan["insert"]))
    for m in plan["insert"]:
        print("  %2d  %.66s" % (m["row"]["n"], m["row"]["artist"] + " - " + m["row"]["title"]))
        print("      show %s" % chipstr(m))
    nochip = [m["row"]["n"] for m in plan["reorder"] + plan["insert"] if not m["chip"][0]]
    if nochip:
        print("\n  %d of %d cards will draw NO show line: %s"
              % (len(nochip), len(rb), ", ".join(map(str, nochip))))
    verb = "trailed after" if a.keep_extras else "dropped, logged passed@%d, carried to the next show" % a.station
    print("\nNOT PLAYED  (%d)   -- %s\n" % (len(plan["drop"]), verb))
    for d in plan["drop"]:
        print("      %-26.26s  %.58s" % (d["artist_name"] or "", d["title"] or ""))

    if not a.apply:
        print("\n" + "-" * 96)
        print("REVIEW ONLY -- nothing written. Re-run with --apply to write to %s." % a.ref)
        return

    # ---- apply --------------------------------------------------------------
    # Walk the export IN PLAY ORDER, not bucket by bucket. `reorder` and `insert`
    # are two lists over one interleaved sequence -- numbering each list in turn
    # puts every kept track first and every new one after it, which is a tracklist
    # in no order at all.
    stmts, sort = [], 10
    walk = ([("keep", m) for m in plan["reorder"]] + [("new", m) for m in plan["insert"]])
    walk.sort(key=lambda km: km[1]["row"]["n"])
    for kind, m in walk:
      if kind == "keep":
        r, d = m["row"], m["db"]
        sets = ["sort = %d" % sort, "title = %s" % lit(r["title"]),
                "artist_name = %s" % lit(r["artist"]), "source = 'rekordbox'"]
        if r["bpm"]:
            sets.append("bpm = %d" % r["bpm"])
        sk, cam = key_fields(r["key"])
        if sk: sets.append("song_key = %s" % lit(sk))
        if cam: sets.append("camelot = %s" % lit(cam))
        if r["duration_ms"]:
            sets.append("duration_ms = %d" % r["duration_ms"])
        if r["genre"]:
            sets.append("genres = array[%s]" % lit(r["genre"]))
        if r["label"]:
            sets.append("label = %s" % lit(r["label"]))
        # The chip is rewritten from ra_artists, or cleared. A chip that has gone
        # stale, or whose artist is no longer credited on the card, is a lie.
        cd, cv = m["chip"]
        sets.append("show_date = %s" % lit(cd))
        sets.append("show_venue = %s" % lit(cv))
        stmts.append("update sc_playlist_tracks set %s where id = '%s';" % (", ".join(sets), d["id"]))
        sort += 10
      else:
        r = m["row"]
        sk, cam = key_fields(r["key"])
        # Synthetic id in the shape manual/Rekordbox tracks use -- sc_track_id is
        # the key of the permanent song memory and stays NOT NULL (migration 102).
        tid = "man_%s_%d" % (re.sub(r"[^a-z0-9]+", "", norm(r["title"]))[:18] or "track", r["n"])
        cols = {"playlist_id": lit(pl["id"]), "sc_track_id": lit(tid), "source": "'rekordbox'",
                "title": lit(r["title"]), "artist_name": lit(r["artist"]),
                "bpm": str(r["bpm"]) if r["bpm"] else "null",
                "song_key": lit(sk), "camelot": lit(cam),
                "duration_ms": str(r["duration_ms"]) if r["duration_ms"] else "null",
                "label": lit(r["label"]), "sort": str(sort),
                "genres": ("array[%s]" % lit(r["genre"])) if r["genre"] else "null",
                "show_date": lit(m["chip"][0]), "show_venue": lit(m["chip"][1])}
        stmts.append("insert into sc_playlist_tracks (%s) values (%s);"
                     % (", ".join(cols), ", ".join(cols[k] for k in cols)))
        stmts.append("insert into sc_song_log (sc_track_id, title, artist_name, duration_ms, source) "
                     "values (%s, %s, %s, %s, 'rekordbox') on conflict (sc_track_id) do nothing;"
                     % (lit(tid), lit(r["title"]), lit(r["artist"]),
                        str(r["duration_ms"]) if r["duration_ms"] else "null"))
        sort += 10
    for d in plan["drop"]:
        if a.keep_extras:
            stmts.append("update sc_playlist_tracks set sort = %d where id = '%s';" % (sort, d["id"]))
            sort += 10
            continue
        # Log passed BEFORE deleting -- that is what carries it into the next
        # station at finalize. Delete without it and the song is just forgotten.
        stmts.append(
            "insert into sc_song_log (sc_track_id, title, artist_name, permalink_url, artwork_url, "
            "duration_ms, source, buy_url, label, passed_playlist_id, passed_station_no, passed_at, updated_at) "
            "values (%s, %s, %s, %s, %s, %s, %s, %s, %s, '%s', %d, now(), now()) "
            "on conflict (sc_track_id) do update set passed_playlist_id = excluded.passed_playlist_id, "
            "passed_station_no = excluded.passed_station_no, passed_at = excluded.passed_at, "
            "updated_at = excluded.updated_at;"
            % (lit(d["sc_track_id"]), lit(d["title"]), lit(d["artist_name"]), lit(d["permalink_url"]),
               lit(d["artwork_url"]), str(d["duration_ms"]) if d["duration_ms"] else "null",
               lit(d["source"] or "soundcloud"), lit(d["buy_url"]), lit(d["label"]),
               pl["id"], a.station))
        stmts.append("delete from sc_playlist_tracks where id = '%s';" % d["id"])

    stmts.append("update sc_playlists set updated_at = now() where id = '%s';" % pl["id"])

    body = "\n".join(stmts)
    print("\n" + "=" * 96)
    print("APPLYING %d statements to %s" % (len(stmts), a.ref))
    print("=" * 96)
    # One transaction, one approval. A half-applied reorder is worse than none.
    sql("begin;\n" + body + "\ncommit;", a.ref)

    after = sql("select sort, artist_name, title, bpm, camelot, source from sc_playlist_tracks "
                "where playlist_id = '%s' order by sort;" % pl["id"], a.ref)
    print("\nSHOW %s now holds %d tracks:\n" % (pl["station_no"], len(after)))
    for i, t in enumerate(after, 1):
        print("  %2d  %-30.30s  %-42.42s  %s %s"
              % (i, t["artist_name"] or "", t["title"] or "", t["camelot"] or "--", t["bpm"] or ""))
    if len(after) != len(rb) and not a.keep_extras:
        print("\n!! expected %d, got %d -- check before publishing." % (len(rb), len(after)))


if __name__ == "__main__":
    main()
