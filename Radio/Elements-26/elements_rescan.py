# Re-pull songs for the Elements artist pool with the songs-not-sets rule applied.
#
# Why this exists rather than re-running elements_tool.py: that script also creates
# ra_artists rows, renames "Day N" -> "Ep N" and re-scopes the four Elements
# episodes to their day's lineup. None of that needs redoing, and re-running it
# would churn episode config to fix a cache problem. This touches ONE thing:
# sc_artist_cache songs / song_count / set_count / is_producer.
#
# Safe to re-run. Pass --dry to see what would change without writing.
import os, json, sys, time, urllib.request
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from elements_sc import fetch_songs, SONG_MAX_MS

ROOT = r"C:\Users\Admin\Documents\Comewith"
DRY = "--dry" in sys.argv
env = {}
for line in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
    line = line.strip()
    if "=" in line and not line.startswith("#"):
        k, v = line.split("=", 1); env[k] = v.strip().strip('"').strip("'")
PAT, REF = env["SBP_PAT"], env.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk")
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}

def sql(q, tries=4):
    # The Management API throws the occasional 502; a single blip must not abandon
    # a 73-row run. Safe to retry: every statement here is idempotent (the SELECT is
    # read-only, the UPDATE sets absolute values rather than incrementing).
    last = None
    for n in range(tries):
        req = urllib.request.Request(f"https://api.supabase.com/v1/projects/{REF}/database/query",
            data=json.dumps({"query": q}).encode(),
            headers={"Authorization": "Bearer " + PAT, "Content-Type": "application/json", **UA}, method="POST")
        try:
            return json.loads(urllib.request.urlopen(req, timeout=40).read().decode() or "null")
        except Exception as e:
            last = e
            if n < tries - 1: time.sleep(1.5 * (n + 1))
    raise last

def esc(s): return str(s).replace("'", "''")

cid = sql("select value from site_content where key='ops.sc_client_id';")[0]["value"]
api = "https://api-v2.soundcloud.com"

# Which rows need a re-pull. Driven off the DATA, not off the Elements lineup, so
# any row with the same problem is repaired regardless of which tool wrote it.
#   over-length song   -> a DJ set stored as a song
#   song_count = 0     -> nothing found; usually a catalogue that lives in
#                         albums/playlists rather than /tracks (LEVEL UP: 46
#                         tracks, 0 served) — the container merge now reaches it
#   --all              -> every Elements artist, e.g. after a rule change like the
#                         publisher-credit check, which needs a re-read to apply
ALL = "--all" in sys.argv
where = ("a.source = 'elements'" if ALL else f"""(
     exists (select 1 from jsonb_array_elements(c.songs) s
             where (s->>'duration_ms')::bigint > {SONG_MAX_MS})
  or coalesce(c.song_count, 0) = 0)""")
rows = sql(f"""
  select distinct c.soundcloud, c.sc_user_id, c.username, c.song_count, c.set_count,
    coalesce(a.name, c.username) as artist_name,
    (select count(*) from jsonb_array_elements(c.songs) s
      where (s->>'duration_ms')::bigint > {SONG_MAX_MS}) as mixes,
    jsonb_array_length(coalesce(c.songs,'[]'::jsonb)) as cached
  from public.sc_artist_cache c
  left join public.ra_artists a
    on lower(rtrim(a.soundcloud,'/')) = lower(rtrim(c.soundcloud,'/'))
  where c.sc_user_id is not null and {where}
  order by mixes desc;""") or []
print(f"{len(rows)} artist(s) to re-pull{' (DRY RUN)' if DRY else ''}{' [--all]' if ALL else ''}\n")

fixed = failed = 0
tot_before = tot_after = tot_dropped = 0
for i, r in enumerate(rows, 1):
    name = r["username"] or r["soundcloud"]
    # Both names, so a collaboration isn't read as someone else's release.
    names = tuple(n for n in (r.get("artist_name"), r.get("username")) if n)
    try:
        songs, sets, dropped = fetch_songs(api, cid, r["sc_user_id"], artist_names=names)
    except Exception as e:
        print(f"  !! {name}: {e}"); failed += 1; continue
    over = [s for s in songs if (s["duration_ms"] or 0) > SONG_MAX_MS]
    assert not over, f"{name}: filter let {len(over)} long track(s) through"
    tot_before += r["mixes"]; tot_after += len(songs); tot_dropped += len(dropped)
    print(f"  {i:>3}/{len(rows)} {str(name)[:26]:<26} was {r['cached']:>2} cached ({r['mixes']} mixes) -> now {len(songs):>2} songs, {sets} sets")
    for title, who in dropped:
        print(f"          dropped, credited to {who}: {str(title)[:50]}")
    if not DRY:
        sj = json.dumps(songs).replace("'", "''")
        sql(f"""update public.sc_artist_cache set songs='{sj}'::jsonb,
              song_count={len(songs)}, set_count={sets}, is_producer={str(len(songs) > 0).lower()},
              scanned_at=now() where soundcloud='{esc(r['soundcloud'])}';""")
        fixed += 1
    time.sleep(0.15)

print(f"\n{'would fix' if DRY else 'updated'} {fixed if not DRY else len(rows)} artist(s), {failed} failed")
print(f"mix rows removed: {tot_before} · wrongly-credited tracks dropped: {tot_dropped}"
      f" · real songs now cached across them: {tot_after}")
if not DRY:
    left = sql(f"""select count(*) as artists from public.sc_artist_cache c
      where exists (select 1 from jsonb_array_elements(c.songs) s
                    where (s->>'duration_ms')::bigint > {SONG_MAX_MS});""")[0]["artists"]
    print(f"artists still holding an over-length song: {left}   {'OK' if left == 0 else 'CHECK THESE'}")
