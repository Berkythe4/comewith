#!/usr/bin/env python3
"""
Apply migration 099 (Come With Radio episodes + listener accounts) to PROD via
the Supabase Management API, with introspection before and verification after.

Run from repo root:  python scripts/apply_099_radio_episodes.py [--check-only]

Needs a VALID SBP_PAT in .env. 401 → refresh at
https://supabase.com/dashboard/account/tokens
"""
import json, os, sys, urllib.request, urllib.error

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

def env():
    e = {}
    for line in open(os.path.join(ROOT, '.env'), encoding='utf-8'):
        line = line.strip()
        if '=' in line and not line.startswith('#'):
            k, v = line.split('=', 1); e[k] = v.strip()
    return e

E = env()
REF, PAT = E['SBP_REF_PROD'], E['SBP_PAT']

def q(sql):
    req = urllib.request.Request(
        'https://api.supabase.com/v1/projects/%s/database/query' % REF,
        data=json.dumps({'query': sql}).encode(),
        headers={'Authorization': 'Bearer ' + PAT, 'Content-Type': 'application/json', 'User-Agent': UA},
        method='POST')
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read().decode() or 'null')
    except urllib.error.HTTPError as ex:
        body = ex.read().decode()
        raise SystemExit('HTTP %s on SQL:\n%s\n--> %s' % (ex.code, sql[:300], body[:800]))

print('== PRE-CHECK (prod %s) ==' % REF)
cols = q("select column_name from information_schema.columns where table_schema='public' and table_name='sc_playlists' order by 1")
colnames = {c['column_name'] for c in cols}
print('sc_playlists columns:', ', '.join(sorted(colnames)))
already = 'status' in colnames and 'slug' in colnames
print('099 already applied?' , already)
rows = q("select id, name, created_at, (select count(*) from sc_playlist_tracks t where t.playlist_id=p.id) tracks from sc_playlists p order by created_at")
for r in rows: print('  station:', r['id'][:8], r['name'], 'tracks=%s' % r['tracks'])
exists = q("select table_name from information_schema.tables where table_schema='public' and table_name in ('listener_playlists','listener_playlist_tracks','listener_station_history','sc_song_log')")
print('new tables already present:', [t['table_name'] for t in exists])

if '--check-only' in sys.argv:
    sys.exit(0)

print('\n== APPLY 099 ==')
sql = open(os.path.join(ROOT, 'supabase', 'migrations', '099_radio_episodes.sql'), encoding='utf-8').read()
q(sql)
print('applied.')

print('\n== VERIFY ==')
cols = {c['column_name'] for c in q("select column_name from information_schema.columns where table_schema='public' and table_name='sc_playlists'")}
need = {'status','slug','station_no','mix_file_path','mix_sc_track_id','mix_sc_track_url','mix_youtube_url','desc_public','desc_sc','published_at','cover_url'}
missing = need - cols
print('sc_playlists new columns OK' if not missing else 'MISSING: %s' % missing)

tabs = [t['table_name'] for t in q("select table_name from information_schema.tables where table_schema='public' and table_name in ('listener_playlists','listener_playlist_tracks','listener_station_history','sc_song_log')")]
print('tables:', tabs)

pols = q("select tablename, policyname from pg_policies where schemaname='public' and tablename in ('listener_playlists','listener_playlist_tracks','listener_station_history','sc_song_log') order by 1")
for p in pols: print('  policy:', p['tablename'], '→', p['policyname'])
rls = q("select relname, relrowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='public' and relname in ('listener_playlists','listener_playlist_tracks','listener_station_history','sc_song_log')")
for r in rls: print('  rls enabled:', r['relname'], r['relrowsecurity'])

b = q("select id, public from storage.buckets where id='radio-mixes'")
print('bucket radio-mixes:', b)
sp = q("select policyname from pg_policies where schemaname='storage' and tablename='objects' and policyname ilike '%radio mixes%'")
print('storage policies:', [p['policyname'] for p in sp])

sn = q("select station_no, name, status from sc_playlists order by station_no")
for r in sn: print('  numbered:', 'EP', r['station_no'], r['name'], r['status'])
bcount = q("select count(*) n from sc_playlists where status='building'")
print('building rows (must be <=1):', bcount[0]['n'])

anon = q("select table_name, privilege_type from information_schema.role_table_grants where grantee='anon' and table_schema='public' and table_name in ('listener_playlists','listener_playlist_tracks','listener_station_history','sc_song_log')")
print('anon grants on new tables (must be []):', anon)

# RLS smoke: an authenticated user (a REAL auth.users id — FK) can create a
# playlist and INSERT..RETURNING a track into it (the 097 lesson). Rolled back.
print('\n== RLS SMOKE (BEGIN..ROLLBACK as authenticated user) ==')
smoke = q("""
begin;
do $$
declare uid uuid; plid uuid; tid uuid;
begin
  select id into uid from auth.users order by created_at limit 1;
  if uid is null then raise notice 'no auth users — skipping'; return; end if;
  perform set_config('request.jwt.claims', json_build_object('sub', uid, 'role', 'authenticated')::text, true);
  perform set_config('role', 'authenticated', true);
  insert into public.listener_playlists (user_id, name) values (uid, 'rls-smoke') returning id into plid;
  insert into public.listener_playlist_tracks (playlist_id, title, permalink_url)
    values (plid, 'smoke track', 'https://soundcloud.com/x/y') returning id into tid;
  raise notice 'RLS smoke OK: playlist % track %', plid, tid;
end $$;
rollback;
""")
print('insert..returning as authenticated owner: OK (rolled back)', smoke)
print('\nDone.')
