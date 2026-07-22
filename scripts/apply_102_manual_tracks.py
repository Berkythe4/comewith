#!/usr/bin/env python3
"""
Apply migration 102 (manual / Rekordbox-sourced station tracks) to PROD via the
Supabase Management API, with introspection before and verification after.

Run from repo root:  python scripts/apply_102_manual_tracks.py [--check-only]

Needs a VALID SBP_PAT in .env. 401 -> refresh at
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

COLS = """
select table_name, column_name, data_type, is_nullable, column_default
from information_schema.columns
where table_schema = 'public'
  and table_name in ('sc_playlist_tracks','sc_song_log')
  and column_name in ('source','buy_url','label','sc_track_id','permalink_url')
order by table_name, column_name;
"""

print('== PRE-CHECK (prod %s) ==' % REF)
pre = q(COLS)
for r in pre:
    print('  %-20s %-14s %-9s null=%s default=%s' % (r['table_name'], r['column_name'], r['data_type'], r['is_nullable'], r['column_default']))
have = {(r['table_name'], r['column_name']) for r in pre}
todo = [c for c in [('sc_playlist_tracks', 'source'), ('sc_playlist_tracks', 'buy_url'), ('sc_playlist_tracks', 'label'),
                    ('sc_song_log', 'source'), ('sc_song_log', 'buy_url'), ('sc_song_log', 'label')] if c not in have]
print('  -> %d column(s) to add: %s' % (len(todo), ', '.join('%s.%s' % c for c in todo) or 'none'))

# permalink_url must already be nullable — a hand-added track can have no public link.
pu = [r for r in pre if r['column_name'] == 'permalink_url']
if pu and pu[0]['is_nullable'] != 'YES':
    raise SystemExit('ABORT: sc_playlist_tracks.permalink_url is NOT NULL on prod — manual tracks need it nullable.')

if '--check-only' in sys.argv:
    sys.exit(0)
if not todo:
    print('\nNothing to do — 102 already applied.')
    sys.exit(0)

sql = open(os.path.join(ROOT, 'supabase', 'migrations', '102_manual_tracks.sql'), encoding='utf-8').read()
print('\n== APPLYING 102 ==')
q(sql)

print('\n== POST-CHECK ==')
for r in q(COLS):
    print('  %-20s %-14s %-9s null=%s default=%s' % (r['table_name'], r['column_name'], r['data_type'], r['is_nullable'], r['column_default']))

print('\n-- source check constraint --')
for r in q("""select conname, pg_get_constraintdef(oid) def from pg_constraint
              where conrelid = 'public.sc_playlist_tracks'::regclass and conname = 'sc_playlist_tracks_source_chk';"""):
    print('  %s: %s' % (r['conname'], r['def']))

print('\n-- RLS still has real policies + anon still blocked --')
for r in q("""select c.relname, c.relrowsecurity, count(p.polname) policies
              from pg_class c left join pg_policy p on p.polrelid = c.oid
              where c.relname in ('sc_playlist_tracks','sc_song_log') group by 1,2;"""):
    print('  %-20s rls=%s policies=%s' % (r['relname'], r['relrowsecurity'], r['policies']))
for r in q("""select table_name, grantee, string_agg(privilege_type, ',' order by privilege_type) privs
              from information_schema.role_table_grants
              where table_schema='public' and table_name in ('sc_playlist_tracks','sc_song_log') and grantee='anon'
              group by 1,2;"""):
    print('  ANON GRANT PRESENT (should be none): %s' % r)

print('\n-- existing rows defaulted correctly --')
for r in q("select source, count(*) n from public.sc_playlist_tracks group by 1 order by 1;"):
    print('  source=%s -> %s rows' % (r['source'], r['n']))

print('\nDone. Commit 102_manual_tracks.sql so tracked history matches prod.')
