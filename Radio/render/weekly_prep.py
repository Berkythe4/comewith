#!/usr/bin/env python3
"""
weekly_prep.py — one read-only command that assembles everything paper-y for an
episode so you're not clicking around gathering it. Writes four files into
Radio/Documents/ (nothing is changed in the DB or on the site):

  EP{N}_cues.csv            — tracklist ready for timestamps + the video render
  EP{N}_youtube.txt         — title, description and a chapters block (paste into
                              the YouTube upload; chapters fill in once the cues
                              have start times)
  EP{N}_buylist.txt         — every track + its Beatport/Bandcamp buy link/price
                              if "Where to buy" has been run (from the DB cache)
  EP{N}_checklist.md        — the release checklist, pre-ticked for what's done

Usage:  python Radio/render/weekly_prep.py            # working station
        python Radio/render/weekly_prep.py --station 1

Read-only: it only SELECTs from prod. Safe to run anytime.
"""
import argparse, csv, json, os, sys, urllib.request, urllib.error
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCS = os.path.join(ROOT, "Radio", "Documents")
RENDER = os.path.join(ROOT, "Radio", "render")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"

def env():
    e = {}
    with open(os.path.join(ROOT, ".env"), encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, v = line.split("=", 1); e[k] = v.strip()
    return e

def q(E, sql):
    req = urllib.request.Request(
        "https://api.supabase.com/v1/projects/%s/database/query" % E["SBP_REF_PROD"],
        data=json.dumps({"query": sql}).encode(),
        headers={"Authorization": "Bearer " + E["SBP_PAT"], "Content-Type": "application/json", "User-Agent": UA},
        method="POST")
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read().decode() or "null")
    except urllib.error.HTTPError as ex:
        raise SystemExit("HTTP %s: %s" % (ex.code, ex.read().decode()[:300]))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--station", type=int)
    ap.add_argument("--week", help="write everything into Radio/Week N/ instead of Documents/ + render/")
    a = ap.parse_args()
    E = env()
    global DOCS, RENDER
    if a.week:
        DOCS = RENDER = os.path.join(ROOT, "Radio", "Week %s" % a.week)
    os.makedirs(DOCS, exist_ok=True)

    where = "p.station_no = %d" % a.station if a.station else "p.status in ('building','testing')"
    st = q(E, """select p.station_no, p.name, p.status, p.drop_date, p.mix_by, p.slug,
                        (p.desc_public is not null) has_desc, p.cover_url,
                        (p.mix_sc_track_url is not null) has_mix, (p.mix_youtube_url is not null) has_yt
                 from sc_playlists p where %s order by p.station_no desc limit 1;""" % where)
    if not st: raise SystemExit("No station found.")
    S = st[0]; N = S["station_no"]
    tr = q(E, """select t.sort, t.artist_name, t.title, t.bpm, t.song_key, t.camelot, t.duration_ms,
                        t.beatport_url, t.beatport_price, t.bandcamp_url, t.source
                 from sc_playlist_tracks t join sc_playlists p on p.id=t.playlist_id
                 where p.station_no=%d order by t.sort;""" % N)
    if not tr: raise SystemExit("Station has no tracks.")

    art = q(E, "select value from site_content where key='ops.radio_artwork';")
    has_art = bool(S["cover_url"] or (art and art[0].get("value")))

    # 1) cues CSV (same as make_cues) ----------------------------------------
    cues = os.path.join(RENDER, "EP%d_cues.csv" % N)
    # Carry forward any start times ALREADY typed in. This used to overwrite the
    # file with blanks and then read it back two steps later, so the chapters were
    # always "(time)" and re-running - which this file tells you to do - silently
    # threw away the timings you had just entered. Keyed on artist+title so a
    # re-ordered tracklist can't shift a time onto the wrong song.
    prior, prior_idx = {}, {}
    if os.path.exists(cues):
        try:
            with open(cues, newline="", encoding="utf-8") as pf:
                for n, row in enumerate(csv.DictReader(pf), 1):
                    if (row.get("start") or "").strip():
                        k = ((row.get("artist") or "").strip().lower(),
                             (row.get("title") or "").strip().lower())
                        prior[k] = row["start"].strip()
                        prior_idx[n] = row["start"].strip()
        except Exception:
            prior, prior_idx = {}, {}
    with open(cues, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["idx", "start", "artist", "title", "bpm", "song_key", "camelot", "duration_ms"])
        for i, r in enumerate(tr, 1):
            # Name match first; fall back to POSITION. Cues built from a Rekordbox
            # export carry the raw deck titles ("Bam Bam (Original Mix)") while the
            # dashboard holds the cleaned ones ("Bam Bam"), so a name-only match
            # dropped 14 of 19 times on EP2. Same length + same order = index is safe.
            kept = prior.get(((r["artist_name"] or "").strip().lower(),
                              (r["title"] or "").strip().lower()), "") or prior_idx.get(i, "")
            w.writerow([i, kept, r["artist_name"] or "", r["title"] or "", r["bpm"] or "",
                        r["song_key"] or "", r["camelot"] or "", r["duration_ms"] or ""])

    # 2) YouTube text (title + description + chapters) ------------------------
    page = "https://comewith.org/radio.html?s=%s" % (S["slug"] or ("come-with-radio-ep%d" % N))
    yt = os.path.join(DOCS, "EP%d_youtube.txt" % N)
    with open(yt, "w", encoding="utf-8") as f:
        f.write("TITLE:\nCome With Radio — EP %d%s\n\n" % (N, (" · %s" % S["name"]) if S["name"] and S["name"] != "Come With Radio" else ""))
        f.write("DESCRIPTION:\n")
        f.write("Come With Radio — EP %d.%s\n" % (N, (" Mixed by %s." % S["mix_by"]) if S["mix_by"] else ""))
        f.write("A monthly mix of artists playing New York soon.\n")
        f.write("Full tracklist, dates & tickets: %s\n\n" % page)
        f.write("CHAPTERS (YouTube turns 0:00-style lines into clickable chapters).\n")
        f.write("Fill EP%d_cues.csv with start times first — then re-run this to bake them in:\n" % N)
        f.write("0:00 %s — %s\n" % (tr[0]["artist_name"] or "", tr[0]["title"] or ""))
        # if the cues already have times, use them
        filled = _read_starts(cues)
        for i, r in enumerate(tr):
            stamp = filled[i] if i < len(filled) and filled[i] else ("(time)" if i else None)
            if stamp is None:  # first line already written
                continue
            f.write("%s %s — %s\n" % (stamp, r["artist_name"] or "", r["title"] or ""))

    # 3) buy list ------------------------------------------------------------
    bl = os.path.join(DOCS, "EP%d_buylist.txt" % N)
    total = 0.0; found = 0
    with open(bl, "w", encoding="utf-8") as f:
        f.write("BUY LIST — Come With Radio EP %d\n\n" % N)
        for i, r in enumerate(tr, 1):
            where_txt = r["beatport_url"] or r["bandcamp_url"] or "(not matched — run 'Where to buy' in the dashboard)"
            price = (" · %s" % r["beatport_price"]) if r["beatport_price"] else ""
            if r["beatport_url"] or r["bandcamp_url"]: found += 1
            if r["beatport_price"]:
                try: total += float(str(r["beatport_price"]).replace("$", "").replace(",", ""))
                except ValueError: pass
            f.write("%2d. %s — %s%s\n    %s\n" % (i, r["artist_name"] or "", r["title"] or "", price, where_txt))
        f.write("\n%d/%d matched to a store. Beatport subtotal: $%.2f\n" % (found, len(tr), total))
        f.write("(To fill in matches + prices, run the dashboard's 🛒 Where to buy, then re-run this.)\n")

    # 4) checklist -----------------------------------------------------------
    def box(done): return "[x]" if done else "[ ]"
    times_done = all(filled) and len(filled) == len(tr)
    cl = os.path.join(DOCS, "EP%d_checklist.md" % N)
    with open(cl, "w", encoding="utf-8") as f:
        f.write("# Come With Radio — EP %d release checklist\n\n" % N)
        f.write("Station: **%s** · status **%s** · drops **%s**\n\n" % (S["name"], S["status"], S["drop_date"] or "—"))
        f.write("## Build\n")
        f.write("- %s %d tracks on the station\n" % (box(len(tr) > 0), len(tr)))
        f.write("- %s Show info filled (↻ Show info)\n" % box(True))
        f.write("- %s Bought the tracks (🛒 Where to buy → /beatport-cart)\n" % box(found > 0))
        f.write("\n## Arrange & record\n")
        f.write("- [ ] Arranged in Rekordbox\n- [ ] Recorded the mix\n")
        f.write("- %s Tapped/exported the track start times (EP%d_cues.csv)\n" % (box(times_done), N))
        f.write("\n## Video (Option 4)\n")
        f.write("- [ ] Rendered EP%d.mp4 (render_episode.py)\n- [ ] Uploaded to YouTube\n" % N)
        f.write("- %s YouTube link pasted in ✎ Details\n" % box(S["has_yt"]))
        f.write("\n## Details & go live\n")
        f.write("- %s Episode name set (not 'Weekly station')\n" % box(S["name"] and S["name"].lower() != "weekly station"))
        f.write("- %s Station artwork / cover set\n" % box(has_art))
        f.write("- %s Descriptions written\n" % box(S["has_desc"]))
        f.write("- %s Mixed-by (DJ) set\n" % box(bool(S["mix_by"])))
        f.write("- [ ] ⇪ To SoundCloud tested (token alive)\n")
        f.write("- %s Mix on SoundCloud\n" % box(S["has_mix"]))
        f.write("- [ ] 🚀 Go live\n- [ ] Homepage teaser shows the right name\n")

    print("Wrote for EP %d:" % N)
    for p in (cues, yt, bl, cl):
        print("  " + os.path.relpath(p, ROOT).replace("\\", "/"))
    print("\nStatus: %d tracks · %d in a store · times %s · artwork %s · desc %s · DJ %s"
          % (len(tr), found, "SET" if times_done else "todo",
             "set" if has_art else "todo", "set" if S["has_desc"] else "todo",
             S["mix_by"] or "todo"))

def _read_starts(cues_path):
    try:
        with open(cues_path, encoding="utf-8-sig") as f:
            return [r.get("start", "").strip() for r in csv.DictReader(f)]
    except Exception:
        return []

if __name__ == "__main__":
    main()
