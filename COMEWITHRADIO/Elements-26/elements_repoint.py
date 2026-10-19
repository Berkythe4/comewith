# Repoint Elements artists whose matched SoundCloud profile is the WRONG one.
#
# Why a script and not a one-off UPDATE: repointing is two writes that must agree
# (ra_artists.soundcloud and the sc_artist_cache row keyed by that URL), and the
# old cache row has to stop being the one the pool reads. Doing it by hand left
# the stale row behind and the artist still looked empty.
#
# ONLY conclusive cases belong in here. The rule (memory: flag suspicious artist
# matches) is that a 0-track / low-follower match is FLAGGED, not silently
# re-pointed — an automated "better guess" that is wrong is worse than a blank,
# because nobody re-checks a profile that looks filled in. Each entry below
# carries the evidence that made it conclusive.
import os, json, re, sys, time, urllib.request
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from elements_sc import fetch_songs

ROOT = r"C:\Users\Admin\Documents\Comewith"
DRY = "--dry" in sys.argv
env = {}
for l in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
    l = l.strip()
    if "=" in l and not l.startswith("#"):
        k, v = l.split("=", 1); env[k] = v.strip().strip('"').strip("'")
PAT, REF = env["SBP_PAT"], env.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk")
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}

def sql(q, tries=4):
    last = None
    for n in range(tries):
        req = urllib.request.Request(f"https://api.supabase.com/v1/projects/{REF}/database/query",
            data=json.dumps({"query": q}).encode(),
            headers={"Authorization": "Bearer " + PAT, "Content-Type": "application/json", **UA}, method="POST")
        try:
            return json.loads(urllib.request.urlopen(req, timeout=60).read().decode() or "null")
        except Exception as e:
            last = e
            if n < tries - 1: time.sleep(1.5 * (n + 1))
    raise last

def esc(s): return str(s).replace("'", "''")
cid = sql("select value from site_content where key='ops.sc_client_id';")[0]["value"]
api = "https://api-v2.soundcloud.com"

# name -> (handle, why it is conclusive)
REPOINT = {
    # /thediis is a 9-follower, 0-track shell. /diisdiis is the same display name,
    # is in Brooklyn, credits every upload to "Diis", and one of them is literally
    # titled "Practice Before Elements". That is the act on the lineup.
    "Diis": ("diisdiis", "uploads credited 'Diis', Brooklyn, track 'Practice Before Elements'"),
    # /raeanne-fissel is a 14-follower, 0-track shell. /mderagon's DISPLAY NAME is
    # "Michael Deragon/Cloud Conductor" — the artist name is in the profile itself.
    "Cloud Conductor": ("mderagon", "profile display name reads 'Michael Deragon/Cloud Conductor'"),
}

def resolve(handle):
    u = f"{api}/resolve?url=https://soundcloud.com/{handle}&client_id={cid}"
    try:
        with urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=20) as r:
            return json.load(r)
    except Exception:
        return None

for name, (handle, why) in REPOINT.items():
    row = sql(f"select ra_id, soundcloud from ra_artists where name='{esc(name)}' and source='elements';")
    if not row:
        print(f"  !! {name}: not in the elements pool"); continue
    old = row[0]["soundcloud"]
    u = resolve(handle)
    if not u or not u.get("id"):
        print(f"  !! {name}: could not resolve @{handle}"); continue
    scu = (u.get("permalink_url") or "").replace("://www.", "://")
    key = scu.strip().lower().rstrip("/").split("?")[0]
    songs, sets, dropped, dupes = fetch_songs(api, cid, u["id"], artist_names=(name, u.get("username") or ""))
    print(f"\n{name}: {old}  ->  {scu}")
    print(f"   why: {why}")
    print(f"   now: {u.get('followers_count') or 0} followers, {u.get('track_count') or 0} uploads "
          f"-> {len(songs)} songs, {sets} sets")
    if DRY:
        continue
    city = (u.get("city") or "").strip() or None
    sj = json.dumps(songs).replace("'", "''")
    sql(f"""insert into sc_artist_cache (soundcloud, sc_user_id, username, avatar_url, city, followers, is_producer, song_count, set_count, sc_track_count, songs, ok, scanned_at)
      values ('{esc(key)}','{u['id']}','{esc(u.get('username') or '')}',{('null' if not u.get('avatar_url') else "'"+esc(u['avatar_url'])+"'")},{('null' if not city else "'"+esc(city)+"'")},{u.get('followers_count') or 0},{str(len(songs)>0).lower()},{len(songs)},{sets},{u.get('track_count') or 0},'{sj}'::jsonb,true,now())
      on conflict (soundcloud) do update set songs=excluded.songs, song_count=excluded.song_count, set_count=excluded.set_count, followers=excluded.followers, city=excluded.city, is_producer=excluded.is_producer, sc_user_id=excluded.sc_user_id, scanned_at=now();""")
    sql(f"""update ra_artists set soundcloud='{esc(scu)}', city={('null' if not city else "'"+esc(city)+"'")},
        follower_count={u.get('followers_count') or 0} where ra_id='{esc(row[0]['ra_id'])}';""")
    # Retire the shell so it can never be read as this artist again.
    sql(f"""delete from sc_artist_cache where lower(rtrim(soundcloud,'/'))=lower(rtrim('{esc(old)}','/'))
        and not exists (select 1 from ra_artists a where lower(rtrim(a.soundcloud,'/'))=lower(rtrim('{esc(old)}','/')));""")
    print("   repointed" + (" (dry run)" if DRY else ""))
print("\nDONE" + (" (DRY RUN — nothing written)" if DRY else ""))
