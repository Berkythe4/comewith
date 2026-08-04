# Put the global SHOW counter (sc_playlists.station_no) back in broadcast order.
#
# WHY: station_no is handed out when an episode is CREATED, not when it airs. The
# four Elements editions were planned after "Come With NYC Radio Ep3" but drop
# two weeks BEFORE it, so the counter reads 3=NYC Ep3 (Aug 20), 4-7=Elements
# (Aug 6-9). A counter of total shows that runs out of order isn't one.
#
# SAFETY:
#   • A published / live episode is NEVER moved. Its number is in its slug, its
#     public page and its played-song history. Only unpublished rows shift.
#   • Numbers are parked in a high range first — station_no is uniquely indexed,
#     so swapping in place would collide mid-update.
#   • sc_song_log.played_station_no / passed_station_no and
#     sc_playlist_tracks.carried_from are remapped in the same transaction; they
#     store the NUMBER, not a foreign key, so they'd silently point at the wrong
#     show otherwise.
#   • Prints the exact inverse mapping so the whole thing can be undone.
#
# Run with --dry (default is dry) to preview; --apply to write.
import os, json, sys, time, urllib.request
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = r"C:\Users\Admin\Documents\Comewith"
APPLY = "--apply" in sys.argv
env = {}
for l in open(os.path.join(ROOT, ".env"), encoding="utf-8"):
    l = l.strip()
    if "=" in l and not l.startswith("#"):
        k, v = l.split("=", 1); env[k] = v.strip().strip('"').strip("'")
PAT, REF = env["SBP_PAT"], env.get("SBP_REF_PROD", "yaytdosxfhcqatmhctzk")
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}

def sql(q, tries=3):
    last = None
    for n in range(tries):
        req = urllib.request.Request(f"https://api.supabase.com/v1/projects/{REF}/database/query",
            data=json.dumps({"query": q}).encode(),
            headers={"Authorization": "Bearer " + PAT, "Content-Type": "application/json", **UA}, method="POST")
        try:
            return json.loads(urllib.request.urlopen(req, timeout=90).read().decode() or "null")
        except urllib.error.HTTPError as e:
            raise SystemExit(f"SQL failed: {e.code} {e.read().decode()[:400]}")
        except Exception as e:
            last = e
            if n < tries - 1: time.sleep(1.5 * (n + 1))
    raise last

rows = sql("""select id::text, station_no, name, status, published,
                     coalesce(drop_date::text,'') as drop_date
              from sc_playlists
              order by (drop_date is null), drop_date, station_no;""")

# Broadcast order: by drop date, undated last (an unscheduled draft is not a show
# yet), ties broken by the existing number so a re-run is stable.
target, n = {}, 0
for r in rows:
    n += 1
    target[r["station_no"]] = n

moved = {old: new for old, new in target.items() if old != new}
frozen = [r for r in rows if (r["published"] or r["status"] == "live") and target[r["station_no"]] != r["station_no"]]
if frozen:
    raise SystemExit("REFUSING: would move a published episode: "
                     + ", ".join(f"{r['name']} ({r['station_no']}->{target[r['station_no']]})" for r in frozen))

print(f"{'APPLY' if APPLY else 'DRY RUN'} — {len(moved)} of {len(rows)} shows move\n")
print(f"  {'now':>3} {'->':^4} {'new':<4} {'drops':<12} {'status':<9} name")
for r in rows:
    old, new = r["station_no"], target[r["station_no"]]
    mark = "  " if old == new else "<-"
    print(f"  {old:>3} {'->':^4} {new:<4} {r['drop_date'] or '(none)':<12} {r['status']:<9} {r['name']} {mark}")

if not moved:
    print("\nAlready in broadcast order — nothing to do."); raise SystemExit(0)

print("\nreferences that get remapped with them:")
for col, tbl in [("played_station_no", "sc_song_log"), ("passed_station_no", "sc_song_log"),
                 ("carried_from", "sc_playlist_tracks")]:
    hits = sql(f"""select {col} as no, count(*) as n from {tbl}
                   where {col} in ({','.join(str(k) for k in moved)}) group by 1 order by 1;""") or []
    for h in hits:
        print(f"  {tbl}.{col} = {h['no']} -> {moved[h['no']]}  ({h['n']} row(s))")
    if not hits:
        print(f"  {tbl}.{col}: none in the moving range")

print("\nto undo:  " + " ".join(f"{new}->{old}" for old, new in moved.items()))

if not APPLY:
    print("\nDRY RUN — nothing written. Re-run with --apply.")
    raise SystemExit(0)

# Park everything high first: station_no is uniquely indexed, so 4->3 while 3
# still exists would collide. One statement batch = one transaction.
OFF = 1000
stmts = [f"update sc_playlists set station_no = station_no + {OFF} where station_no is not null;"]
for old, new in target.items():
    stmts.append(f"update sc_playlists set station_no = {new} where station_no = {old + OFF};")
for col, tbl in [("played_station_no", "sc_song_log"), ("passed_station_no", "sc_song_log"),
                 ("carried_from", "sc_playlist_tracks")]:
    cases = " ".join(f"when {old} then {new}" for old, new in moved.items())
    stmts.append(f"update {tbl} set {col} = case {col} {cases} else {col} end "
                 f"where {col} in ({','.join(str(k) for k in moved)});")
sql("begin;\n" + "\n".join(stmts) + "\ncommit;")

after = sql("""select station_no, name, coalesce(drop_date::text,'') as d from sc_playlists
               order by station_no;""")
print("\nafter:")
for r in after:
    print(f"  SHOW {r['station_no']:<3} {r['d'] or '(none)':<12} {r['name']}")
bad = sql("""select count(*) as n from (
               select station_no, row_number() over (order by (drop_date is null), drop_date) rn
               from sc_playlists) x where station_no <> rn;""")[0]["n"]
print(f"\nout-of-order rows remaining: {bad}   {'OK' if bad == 0 else 'CHECK'}")
