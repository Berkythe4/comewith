# Scope the THURSDAY edition (Ep1) to every producer on the Elements bill.
#
# Ep2/3/4 each get their own day's lineup, which is right — those are the big
# nights and the DJ is playing to a crowd that just saw, or is about to see,
# those acts. Thursday is the early slot: a much smaller bill (10 artists), and
# the audience is arriving for the whole weekend, not for Thursday. So Ep1 gets
# the WHOLE festival to dig through.
#
# PRODUCERS only, festival-wide. A DJ-only profile contributes nothing to a radio
# episode — you cannot play a 90-minute live set as a track — so the ~19 mix-only
# acts are left out rather than padding the crate with dead ends. Every artist
# carries a day tag so Thursday's own bill is still one click away in dj.html.
#
# Reads the day map from the lineup literals in elements_tool.py / elements_disco.py
# rather than re-declaring it — a third copy is a third thing to drift.
#
# It deliberately does NOT derive days from the episodes' own dj_search_params.
# That was the first cut and it was not re-runnable: this script writes all 139
# artists into Ep1, so a second run read them all back as Ep1's day and collapsed
# Fri/Sat/Sun into "Thu". Never take your own output as input.
#
# Safe to re-run. Pass --dry to see the result without writing.
import ast, os, json, sys, time, urllib.request
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

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

EDITION = "Come With Elements Radio"
DISCO_LABEL = "Disco Den"

HERE = os.path.dirname(os.path.abspath(__file__))

def literal(filename, varname):
    """Pull a module-level literal out of a sibling script without running it.

    elements_tool.py and elements_disco.py do their work at import time, so they
    cannot simply be imported for their lineup constants.
    """
    tree = ast.parse(open(os.path.join(HERE, filename), encoding="utf-8").read())
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == varname for t in node.targets):
            return ast.literal_eval(node.value)
    sys.exit(f"could not find {varname} in {filename}")

LINEUP = literal("elements_tool.py", "LINEUP")     # {day: [names]}
DISCO = literal("elements_disco.py", "DISCO")      # [names]

# name -> day it plays. Main lineup wins over Disco Den, and Thursday wins over
# the rest of the weekend: acts booked twice (Crankdat and Mersiv play Thu AND
# Fri) belong on Keith's night in Keith's crate.
day_of = {}
for day in ("Thu", "Fri", "Sat", "Sun"):
    for nm in LINEUP.get(day, []):
        day_of.setdefault(nm, day)
disco_only = [nm for nm in dict.fromkeys(DISCO) if nm not in day_of]
for nm in disco_only:
    day_of[nm] = DISCO_LABEL

# Every Elements artist that actually has songs. is_producer is the scan's own
# verdict; song_count is checked too so a stale flag can't smuggle in an empty.
rows = sql("""
  select a.name, coalesce(a.follower_count,0) as followers, coalesce(c.song_count,0) as songs
  from ra_artists a
  join sc_artist_cache c on lower(rtrim(c.soundcloud,'/'))=lower(rtrim(a.soundcloud,'/'))
  where a.source='elements' and c.is_producer and coalesce(c.song_count,0) > 0;""") or []
songs_of = {r["name"]: r["songs"] for r in rows}
fol_of = {r["name"]: r["followers"] for r in rows}

ORDER = ["Thu", "Fri", "Sat", "Sun", DISCO_LABEL]
def sort_key(nm):
    d = day_of.get(nm)
    return (ORDER.index(d) if d in ORDER else len(ORDER), -fol_of.get(nm, 0), nm.lower())

# Thursday's own bill leads, then the rest of the weekend, then Disco Den — the
# order dj-station preserves, so the DJ's own night is what loads first.
producers = sorted(songs_of, key=sort_key)
if not producers:
    sys.exit("no producers found in the Elements pool — run elements_rescan.py first")

params = {
    "pool": "elements",
    "day": "Thu",
    "scope": "all-producers",       # dj-station -> scope.reach, changes the blurb
    "artists": producers,
    "disco": [n for n in producers if day_of.get(n) == DISCO_LABEL],
    "day_of": {n: day_of[n] for n in producers if n in day_of},
    "weeks": 4,
}

by_day = {}
for n in producers:
    by_day.setdefault(day_of.get(n, "?"), []).append(n)
print(f"Ep1 (Thu) -> {len(producers)} producers, {sum(songs_of.values())} songs total")
for d in ORDER:
    if d in by_day:
        print(f"   {d:<10} {len(by_day[d]):>3} producers · {sum(songs_of[n] for n in by_day[d]):>4} songs")
missing = [n for n in songs_of if n not in day_of]
if missing:
    print(f"   (no day tag: {', '.join(sorted(missing))})")
skipped = sql("""
  select a.name from ra_artists a
  join sc_artist_cache c on lower(rtrim(c.soundcloud,'/'))=lower(rtrim(a.soundcloud,'/'))
  where a.source='elements' and coalesce(c.song_count,0)=0 order by a.name;""") or []
print(f"\nleft out — {len(skipped)} act(s) with no songs on SoundCloud (DJ sets only):")
print("   " + ", ".join(r["name"] for r in skipped))

if DRY:
    print("\nDRY RUN — nothing written")
else:
    pj = json.dumps(params, ensure_ascii=False).replace("'", "''")
    sql(f"""update sc_playlists set dj_search_params='{pj}'::jsonb
            where edition_name='{EDITION}' and edition_seq=1;""")
    print("\nEp1 re-scoped.")
