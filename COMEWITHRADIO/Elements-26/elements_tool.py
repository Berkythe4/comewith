# Build the Come With Elements Radio pool: match each festival artist to
# SoundCloud, pull their songs into sc_artist_cache, create source='elements'
# ra_artists rows, and scope the 4 Elements episodes (Ep1-4 = Thu/Fri/Sat/Sun)
# to their day's lineup. Also renames Day N -> Ep N.
import os, json, re, urllib.request, urllib.parse, unicodedata, time, sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from elements_sc import fetch_songs, SONG_MAX_MS   # the songs-not-sets rule
ROOT = r"C:\Users\Admin\Documents\Comewith"
env = {}
for line in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
    line = line.strip()
    if "=" in line and not line.startswith("#"):
        k, v = line.split("=", 1); env[k] = v.strip()
PAT, REF = env["SBP_PAT"], env["SBP_REF_PROD"]
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}
def sql(q):
    req = urllib.request.Request(f"https://api.supabase.com/v1/projects/{REF}/database/query",
        data=json.dumps({"query": q}).encode(),
        headers={"Authorization": "Bearer " + PAT, "Content-Type": "application/json", **UA}, method="POST")
    return json.loads(urllib.request.urlopen(req).read().decode() or "null")
def esc(s): return str(s).replace("'", "''")
cid = sql("select value from site_content where key='ops.sc_client_id';")[0]["value"]
api = "https://api-v2.soundcloud.com"

# ---- the lineup, per festival day (from the 4 images) --------------------------
LINEUP = {
 "Thu": ["Sunni D","Jack What?","Ardalan","San Pacho","Mersiv","Chris Lorenzo","Wooli","Crankdat","Saka","Austeria"],
 "Fri": ["Chris Lake","Above & Beyond","It's Murph","Jigitz","Kaleena Zanders","Dirtwire","Excision","Crankdat","Atliens","Zingara",
   "Ganja White Night","Mersiv","Dice Man","Mikayli","Mickman","Big Gigantic","Gorillat","Ivy Lab","Effin","Wonkywilla","Subfeels",
   "Rudashi","DJ Shakey","Illexxandra","Kattana","Jelly Bean","Boys Noize","MCR-T","KETTAMA","X Club","Dreya V","Bardo","Ammo Amor","Gavin Blac","Fable"],
 "Sat": ["Subtronics","Of The Trees","Level Up","Clozee","Opiuo","Probcause","Skysia","Cloonee","Matroda","Ayybo","Westend","Louis The Child",
   "9b49","Earth Signs","Alec","Ecamp","Svdden Death","Ray Volpe","Hol!","Hedex","Sippy","Nikita The Wicked","MLE","MPH","Henry Pope","DCAL",
   "Biscits","Linska","Discip","Roddy Lima","Papyon","A-Trak","Sirens","Refrakt","Eric Remy"],
 "Sun": ["Porter Robinson","LSDream","Daily Bread","Know Good","Tractorbeam","The Motet","Marvel Years","I Hate Models","Charlotte de Witte",
   "Tiga","Azzecca","Griz","Josh Teed","Sub Focus","YDG","Chyl","Koopmusik","Lumasi","Cloud Conductor","Thought Process","Zejibo","Brainrack",
   "Pnther","Lightcode","Walker & Royce","Acraze","Will Clarke","Jackie Hollander","Golden Pony","Dr. Chaii","Barz","Luna Mar"],
}
def norm(s):
    s = unicodedata.normalize("NFKD", s or ""); s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())
SUF = ["official","music","sounds","records","recordings","audio","real","the","dj"]
def strip_affix(n, q):
    for s in SUF:
        if len(n) > len(q) and n.startswith(s) and n[len(s):] == q: return q
        if len(n) > len(q) and n.endswith(s) and n[:-len(s)] == q: return q
    return n
# Profiles the name matcher CANNOT reach, because the artist appended a tagline to
# their SoundCloud display name. "KETTAMA (G-TOWN FOREVER)" normalises to
# kettamagtownforever, so an exact-name match instead found a 9-follower, 0-track
# impostor at /kettama and that lineup act had an empty crate (caught 2026-07-30).
# Value = the handle to pin. We reference him as KETTAMA — the tagline is not his name.
HANDLE_PINS = {"kettama": "kettamabro"}
def match(name):
    q = norm(name)
    pin = HANDLE_PINS.get(q)
    if pin:
        u = f"{api}/resolve?url=https://soundcloud.com/{pin}&client_id={cid}"
        try: return json.load(urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=15))
        except Exception: pass          # fall through to the normal search
    u = f"{api}/search/users?q={urllib.parse.quote(name)}&limit=8&client_id={cid}"
    try: js = json.load(urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=15))
    except Exception: return None
    ex = [c for c in js.get("collection", []) if q in [strip_affix(norm(c.get(f) or ""), q) for f in ("permalink","username","full_name")]]
    ex.sort(key=lambda c: (not c.get("verified"), -(c.get("followers_count") or 0), -(c.get("track_count") or 0)))
    return ex[0] if ex else None
# SONGS, NOT DJ SETS — one shared definition, in elements_sc.py, matching the
# contract sc-enrich enforces (45s <= d <= 15 min). This used to be inline here
# with only the 45-second floor and no ceiling, which put 251 multi-hour sets in
# sc_artist_cache as "songs" on 2026-07-28.
def tracks(uid, want=None, names=()):
    # want=None -> the artist's whole catalogue. A cap here is invisible in the
    # crate: 15 songs looks complete whether they have 15 or 150.
    songs, sets, dropped, dupes = fetch_songs(api, cid, uid, artist_names=names, want=want)
    for title, who in dropped:
        print(f"      skipped (credited to {who}): {title}")
    for lost, kept in dupes:
        print(f"      duplicate, kept the higher-clout upload: {lost!r} -> {kept!r}")
    return songs, sets

# unique names, remember all days each belongs to
name_days = {}
for day, names in LINEUP.items():
    for nm in names: name_days.setdefault(nm, set()).add(day)
print(f"{len(name_days)} unique artists to process")

matched = {}   # name -> soundcloud url
done = hit = 0
for nm in name_days:
    done += 1
    m = match(nm); time.sleep(0.15)
    if not m: print(f"  · {nm}: no match"); continue
    scu = (m.get("permalink_url") or "").replace("://www.", "://")
    if not scu: continue
    # Pass every name this profile goes by so a collaboration isn't mistaken for
    # someone else's release (nm = lineup name, username = SoundCloud handle).
    songs, set_count = tracks(m["id"], names=(nm, m.get("username") or "")); time.sleep(0.15)
    # A booked festival artist with no tracks or a handful of followers is almost
    # always a WRONG match, not an artist without music. Say so; do not bury it.
    if not (m.get("track_count") or 0) or (m.get("followers_count") or 0) < 500:
        print(f"  ?? CHECK {nm}: matched {m.get('username')!r} "
              f"({m.get('followers_count') or 0} followers, {m.get('track_count') or 0} tracks, "
              f"{len(songs)} songs) -> {scu}")
    matched[nm] = scu; hit += 1
    key_norm = scu.strip().lower().replace("://www.", "://").rstrip("/").split("?")[0]
    city = (m.get("city") or "").replace("'", "''") or None
    # upsert sc_artist_cache
    sj = json.dumps(songs).replace("'", "''")
    sql(f"""insert into sc_artist_cache (soundcloud, sc_user_id, username, avatar_url, city, followers, is_producer, song_count, set_count, sc_track_count, songs, ok, scanned_at)
      values ('{esc(key_norm)}','{m['id']}','{esc(m.get('username') or '')}',{('null' if not m.get('avatar_url') else "'"+esc(m['avatar_url'])+"'")},{('null' if not city else "'"+city+"'")},{m.get('followers_count') or 0},{str(len(songs)>0).lower()},{len(songs)},{set_count},{m.get('track_count') or 0},'{sj}'::jsonb,true,now())
      on conflict (soundcloud) do update set songs=excluded.songs, song_count=excluded.song_count, set_count=excluded.set_count, followers=excluded.followers, city=excluded.city, is_producer=excluded.is_producer, scanned_at=now();""")
    # upsert ra_artists (source='elements'; no next_event_date so it stays out of the NYC scope)
    slug = re.sub(r'[^a-z0-9]+','-', nm.lower()).strip('-')
    gj = None
    sql(f"""insert into ra_artists (ra_id, name, soundcloud, city, follower_count, source)
      values ('elem_{esc(slug)}','{esc(nm)}','{esc(scu)}',{('null' if not city else "'"+city+"'")},{m.get('followers_count') or 0},'elements')
      on conflict (ra_id) do update set soundcloud=excluded.soundcloud, city=excluded.city, follower_count=excluded.follower_count, source='elements';""")
    if done % 20 == 0: print(f"  {done}/{len(name_days)} · {hit} matched")

print(f"\nMatched {hit}/{len(name_days)}")
# ---- scope + rename the 4 episodes -------------------------------------------
DAY_FOR_SEQ = {1:"Thu",2:"Fri",3:"Sat",4:"Sun"}
for seq, day in DAY_FOR_SEQ.items():
    names = [nm for nm in LINEUP[day] if nm in matched]
    params = {"pool": "elements", "day": day, "artists": names, "weeks": 4}
    pj = json.dumps(params).replace("'", "''")
    sql(f"""update sc_playlists set name='Come With Elements Radio — Ep{seq}', dj_search_params='{pj}'::jsonb
        where edition_name='Come With Elements Radio' and edition_seq={seq};""")
    print(f"  Ep{seq} ({day}): scoped to {len(names)} artists, renamed")
print("DONE")
