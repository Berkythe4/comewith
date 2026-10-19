#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""EVERY upcoming NYC show for every artist credited on each track of an export.

apply_station_from_rekordbox.py writes ONE chip per track (the soonest show of
the first credited name it can match) because sc_playlist_tracks has one
show_date / show_venue. A card can say more than that: "Tijuana (Vintage Culture
Remix)" is Bedouin's song AND Vintage Culture's remix, and both are playing NYC.
This lists all of them, with who is playing, so a card can name each one.

    python Radio/render/all_shows.py --txt "Radio/Episode 7/CWR_Ep7_V2.txt" \
        --from 2026-10-22 --out "Radio/Episode 7/EP7_shows.json"

Read-only. Same crediting (credits_of: split artists + remixer) and the same
lineup index (ra_events.lineup, never ra_artists.next_event_date) as the sync
tool, so the card and the station can't disagree about who counts. The same
room on the same date from two feeds (RA + DICE, or a manual row beside RA's)
collapses to one show.
"""
import argparse, io, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from apply_station_from_rekordbox import (read_rekordbox, credits_of, chip_keys,
                                          norm, sql, lit, PROD)

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--txt", required=True)
    ap.add_argument("--from", dest="since", required=True, metavar="YYYY-MM-DD",
                    help="the drop date -- a show before it has happened when a listener hears it")
    ap.add_argument("--out", required=True)
    ap.add_argument("--ref", default=PROD)
    a = ap.parse_args()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", a.since):
        raise SystemExit("--from wants YYYY-MM-DD")

    rows = read_rekordbox(a.txt)
    key_name = {}                       # lookup key -> the credited name as printed
    for r in rows:
        for n in credits_of(r):
            for k in chip_keys(n):
                key_name.setdefault(k, n)
    vals = ", ".join("(%s)" % lit(k) for k in sorted(key_name))
    q = ("with want(k) as (values %s) "
         "select distinct w.k, l->>'name' as lname, e.event_date::text as d, coalesce(v.name, e.venue_name) as venue, "
         "       e.venue_id::text as vid, e.title, e.source, e.content_url as url "
         "  from ra_events e left join venues v on v.id = e.venue_id "
         " cross join lateral jsonb_array_elements(coalesce(e.lineup,'[]'::jsonb)) l "
         "  join want w on w.k = lower(regexp_replace(l->>'name','[^a-zA-Z0-9]','','g')) "
         " where e.event_date >= '%s'::date and coalesce(v.name, e.venue_name) is not null "
         " order by 2, 3;" % (vals, a.since))
    hits = sql(q, a.ref)

    out = []
    for r in rows:
        keys = {}
        for n in credits_of(r):
            for k in chip_keys(n):
                keys.setdefault(k, n)
        shows = {}
        for h in hits:
            if h["k"] not in keys:
                continue
            # The name as the LISTING prints it -- the credit parse can leave
            # debris ("Vintage Culture &" out of "Vintage Culture & Kiko Franco Remix").
            who = h["lname"] or keys[h["k"]]
            # The SQL key deletes non-ASCII, so "Âme" keys as "me" -- the same as
            # "&ME", a different act. Its accent-FOLDING key ("ame") must be one of
            # the credited keys too, or it is somebody else.
            if norm(who).replace(" ", "") not in keys:
                continue
            # one show per (date, room); venue_id when resolved, else the name
            sk = (h["d"], h["vid"] or h["venue"].strip().lower())
            s = shows.setdefault(sk, {"date": h["d"], "venue": h["venue"], "title": h["title"],
                                      "url": h["url"], "sources": [], "artists": []})
            if h["source"] not in s["sources"]:
                s["sources"].append(h["source"])
            if who.lower() not in [x.lower() for x in s["artists"]]:
                s["artists"].append(who)
            if h["source"] == "ra" and h["url"]:
                s["url"] = h["url"]
        # Two rooms, same night, same artists is one show the venue table has not
        # merged yet ("ONE 77" / "One77 Brooklyn", "SILO" / "SILO Brooklyn"). Collapse
        # it for the CARD only, keeping RA's spelling; the database ruling stays a
        # person's call in v_venue_name_review (never auto-merge a fuzzy match).
        sl, seen = [], {}
        for s in sorted(shows.values(), key=lambda s: (s["date"], "ra" not in s["sources"], s["venue"])):
            dk = (s["date"], tuple(sorted(x.lower() for x in s["artists"])))
            if dk in seen:
                keep = seen[dk]
                keep.setdefault("also_listed_as", []).append(s["venue"])
                keep["sources"] += [x for x in s["sources"] if x not in keep["sources"]]
                continue
            seen[dk] = s; sl.append(s)
        out.append({"n": len(out) + 1, "artist": r["artist"], "title": r["title"],
                    "credited": list(dict.fromkeys(keys.values())), "shows": sl})

    io.open(a.out, "w", encoding="utf-8").write(json.dumps(out, indent=1, ensure_ascii=False))
    for t in out:
        print("%2d  %s - %s" % (t["n"], t["artist"], t["title"]))
        if not t["shows"]:
            print("      (no upcoming NYC show)")
        for s in t["shows"]:
            print("      %s  %-24s %-30s %s" % (s["date"], s["venue"][:24], ", ".join(s["artists"])[:30],
                                                (s["title"] or "")[:50]))
    print("\nwrote", a.out)


if __name__ == "__main__":
    main()
