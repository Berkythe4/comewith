# Add the Disco Den stage (weekend, Fri/Sat/Sun = Ep2/3/4) to the Elements pool,
# tagged so DJs can filter them out from the main lineup.
import os, json, re, urllib.request, urllib.parse, unicodedata, time, sys
sys.stdout.reconfigure(encoding="utf-8")
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

DISCO = ["Ade","Adyy Love","Andrew Highland","Asmot","Brella","Cookies","DCAL","Devin Kroes","Diis","DJ Bombay","DJ Dad",
 "Dr.EZ","Elkind","Family Funktion","Flash Gea","Footwrk","Freq","Funky Pickles","Giolibri","Gjoka Drejaj","Gryph","Izzy Insane",
 "Jack What?","Jaesyun","J.Gill","Jon Pantofel","Just James","Lasad","Little New Yorkers","Los","Mike Kerrigan","Mind Matter",
 "Modnyy","Notdeadyet","Odd Arcana","Oscar N","Otternonsense","PR3ME","Proactive Panic","Riptyde","Rive & Bianca","Rvnskll",
 "Samk","Sirens","Sketchy Pete","Summer Almeida","Szymon Dudzik","Takes Two","Thebusiness","Two Lights","Wan","Willdabeast"]

def norm(s):
    s = unicodedata.normalize("NFKD", s or ""); s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())
SUF = ["official","music","sounds","records","recordings","audio","real","the","dj"]
def strip_affix(n, q):
    for s in SUF:
        if len(n) > len(q) and n.startswith(s) and n[len(s):] == q: return q
        if len(n) > len(q) and n.endswith(s) and n[:-len(s)] == q: return q
    return n
def match(name):
    q = norm(name)
    u = f"{api}/search/users?q={urllib.parse.quote(name)}&limit=8&client_id={cid}"
    try: js = json.load(urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=15))
    except Exception: return None
    ex = [c for c in js.get("collection", []) if q in [strip_affix(norm(c.get(f) or ""), q) for f in ("permalink","username","full_name")]]
    ex.sort(key=lambda c: (not c.get("verified"), -(c.get("followers_count") or 0), -(c.get("track_count") or 0)))
    return ex[0] if ex else None
def tracks(uid):
    u = f"{api}/users/{uid}/tracks?limit=20&client_id={cid}"
    try: js = json.load(urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=15))
    except Exception: return []
    out = []
    for t in js.get("collection", []):
        if t.get("kind") != "track" or (t.get("duration") or 0) < 45000: continue
        out.append({"sc_track_id": str(t["id"]), "title": t.get("title"), "permalink_url": t.get("permalink_url"),
                    "duration_ms": t.get("duration"), "playback_count": t.get("playback_count") or 0,
                    "created_at": t.get("created_at"), "artwork_url": t.get("artwork_url")})
        if len(out) >= 15: break
    return out

matched = {}
for i, nm in enumerate(dict.fromkeys(DISCO), 1):
    m = match(nm); time.sleep(0.15)
    if not m: print(f"  · {nm}: no match"); continue
    scu = (m.get("permalink_url") or "").replace("://www.", "://")
    if not scu: continue
    songs = tracks(m["id"]); time.sleep(0.15)
    matched[nm] = scu
    key = scu.strip().lower().replace("://www.", "://").rstrip("/").split("?")[0]
    city = (m.get("city") or "").replace("'", "''") or None
    sj = json.dumps(songs).replace("'", "''")
    sql(f"""insert into sc_artist_cache (soundcloud, sc_user_id, username, avatar_url, city, followers, is_producer, song_count, sc_track_count, songs, ok, scanned_at)
      values ('{esc(key)}','{m['id']}','{esc(m.get('username') or '')}',{('null' if not m.get('avatar_url') else "'"+esc(m['avatar_url'])+"'")},{('null' if not city else "'"+city+"'")},{m.get('followers_count') or 0},{str(len(songs)>0).lower()},{len(songs)},{m.get('track_count') or 0},'{sj}'::jsonb,true,now())
      on conflict (soundcloud) do update set songs=excluded.songs, song_count=excluded.song_count, followers=excluded.followers, city=excluded.city, is_producer=excluded.is_producer, scanned_at=now();""")
    slug = re.sub(r'[^a-z0-9]+','-', nm.lower()).strip('-')
    sql(f"""insert into ra_artists (ra_id, name, soundcloud, city, follower_count, source)
      values ('elem_{esc(slug)}','{esc(nm)}','{esc(scu)}',{('null' if not city else "'"+city+"'")},{m.get('followers_count') or 0},'elements')
      on conflict (ra_id) do update set soundcloud=excluded.soundcloud, city=excluded.city, follower_count=excluded.follower_count, source='elements';""")
    if i % 15 == 0: print(f"  {i}/{len(DISCO)} · {len(matched)} matched")

disco_matched = [n for n in DISCO if n in matched]
print(f"\nDisco Den matched: {len(disco_matched)}/{len(DISCO)}")

# Re-scope the weekend episodes: main first, then Disco-Den-only, tagged in `disco`.
for seq in (2, 3, 4):
    row = sql(f"select dj_search_params as p from sc_playlists where edition_name='Come With Elements Radio' and edition_seq={seq};")[0]["p"]
    main = row.get("artists", [])
    disco_only = [n for n in disco_matched if n not in main]
    params = {"pool": "elements", "day": row.get("day"), "artists": main + disco_only, "disco": disco_only, "weeks": 4}
    pj = json.dumps(params).replace("'", "''")
    sql(f"update sc_playlists set dj_search_params='{pj}'::jsonb where edition_name='Come With Elements Radio' and edition_seq={seq};")
    print(f"  Ep{seq}: {len(main)} main + {len(disco_only)} disco = {len(main)+len(disco_only)} total")
print("DONE")
