"""Find the SoundCloud track for radio songs whose own id is synthetic.

Songs that came in from a Rekordbox cue sheet ('cue_…') or were added by hand
('man_…') have no SoundCloud id, so they could not link out or go into a
SoundCloud playlist - even though most of them ARE on SoundCloud. This finds them
and writes sc_playlist_tracks.sc_link_id / sc_link_url / sc_link_how (236).
sc_track_id is never touched: it is the song's identity (song log, dedupe,
carry-over all key on it).

Two passes, most trusted first:
  cache   the credited artist's OWN uploads, already in sc_artist_cache (the scan
          read them). No API calls.
  search  SoundCloud search through sc-connect `search_tracks` (Keith's stored
          connection), accepted only when the uploader IS a credited artist, or the
          title names a credited artist ("Artist - Title", a label upload).

Matching is deliberately strict (CLAUDE.md "store matching is adversarial"):
  * the core title must be EQUAL after normalising (case, accents, &/and,
    punctuation, feat. credits, "Artist - " prefixes) - never a substring;
  * the VERSION must be equal: "(X Remix)" only matches "(X Remix)". Standard
    qualifiers (Original / Extended / Radio Edit / Club Mix) are not versions.
No link beats the wrong link: a miss is reported, never guessed.

    SBP_REF=yaytdosxfhcqatmhctzk python scripts/link_radio_tracks_soundcloud.py          # dry run
    SBP_REF=yaytdosxfhcqatmhctzk python scripts/link_radio_tracks_soundcloud.py --write
    ... --no-search      cache pass only
"""
import json
import os
import re
import subprocess
import sys
import unicodedata
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REF = os.environ.get('SBP_REF') or ''
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
WRITE = '--write' in sys.argv
SEARCH = '--no-search' not in sys.argv
if REF != 'yaytdosxfhcqatmhctzk':
    sys.exit('pass SBP_REF=yaytdosxfhcqatmhctzk explicitly')

SPLIT = re.compile(r',|&|\bfeat\.?\b|\bft\.?\b|\bx\b|\bvs\.?\b', re.I)   # = RADIO_SPLIT_RE
VERSION_WORDS = re.compile(r'\b(remix|flip|edit|vip|bootleg|rework|re-?work|dub|mix|version|refix|remake|cover)\b', re.I)
QUALIFIERS = {'original', 'extended', 'radio', 'club', 'mix', 'edit', 'version', 'the', 'a'}


def db(sql):
    out = subprocess.run([sys.executable, 'db.py', '-'], input=sql, capture_output=True, text=True,
                         cwd=ROOT, env={**os.environ, 'SBP_REF': REF}, encoding='utf-8')
    if out.returncode:
        sys.exit('db.py failed: ' + out.stdout[-400:] + out.stderr[-400:])
    t = out.stdout
    return json.loads(t[t.index('['):])


def fold(s):
    s = unicodedata.normalize('NFKD', s or '')
    s = ''.join(c for c in s if not unicodedata.combining(c)).lower()
    s = s.replace('&', ' and ').replace('$', 's')
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def parts(artist):
    return [p.strip() for p in SPLIT.split(artist or '') if p.strip()]


def split_title(title, artists):
    """-> (core, version-key). Strips feat. credits, an 'Artist - ' prefix that names a
    credited artist (or any prefix, for search results), and bracketed qualifiers."""
    t = title or ''
    t = re.sub(r'\s*[\(\[]\s*(feat|ft|featuring|with)\b[^\)\]]*[\)\]]', '', t, flags=re.I)
    t = re.sub(r'\s+(feat|ft|featuring)\.?\s+[^\(\[\-]*', ' ', t, flags=re.I)
    version = []
    for m in re.findall(r'[\(\[]([^\)\]]*)[\)\]]', t):
        if VERSION_WORDS.search(m):
            version.append(m)
    t = re.sub(r'[\(\[][^\)\]]*[\)\]]', ' ', t)
    if ' - ' in t:
        left, right = t.split(' - ', 1)
        if VERSION_WORDS.search(right) and not VERSION_WORDS.search(left) and len(fold(right).split()) <= 5 and any(fold(a) in fold(right) for a in artists):
            version.append(right); t = left            # "Title - X Remix"
        else:
            t = right                                  # "Artist - Title"
    vkey = sorted(w for w in fold(' '.join(version)).split() if w not in QUALIFIERS and not VERSION_WORDS.fullmatch(w))
    return fold(t), ' '.join(vkey)


def same(our, cand_title, artists):
    a = split_title(our, artists)
    b = split_title(cand_title, artists)
    return bool(a[0]) and a == b


# ── what needs a link ──────────────────────────────────────────────────────────
tracks = db("""
select t.id, t.sc_track_id, t.title, t.artist_name, p.station_no
from sc_playlist_tracks t join sc_playlists p on p.id = t.playlist_id
where t.sc_track_id !~ '^\\d+$' and t.sc_link_id is null
order by p.station_no, t.sort;""")
print('[need] %d songs with a synthetic id and no SoundCloud link' % len(tracks))
if not tracks:
    sys.exit(0)

# ── pass 1: the credited artists' own uploads in the scan cache ────────────────
names = sorted({p.lower() for t in tracks for p in parts(t['artist_name'])})
lit = ','.join("'" + n.replace("'", "''") + "'" for n in names)
cache = db("""
with links as (
  select distinct lower(btrim(a.name)) nm,
         regexp_replace(lower(btrim(a.soundcloud)), '^https?://(www\\.|m\\.)?', 'https://') sc
  from ra_artists a where a.soundcloud is not null and lower(btrim(a.name)) in (%s))
select c.username, l.nm credited, s->>'sc_track_id' sid, s->>'title' title, s->>'permalink_url' url,
       coalesce((s->>'playback_count')::numeric, 0) plays
from sc_artist_cache c
left join links l on rtrim(l.sc, '/') = rtrim(c.soundcloud, '/')
cross join lateral jsonb_array_elements(case when jsonb_typeof(c.songs) = 'array' then c.songs else '[]'::jsonb end) s
where l.nm is not null or lower(btrim(c.username)) in (%s);""" % (lit, lit))
by_artist = {}
for s in cache:
    for k in {(s['credited'] or '').lower(), fold(s['username'] or '')} - {''}:
        by_artist.setdefault(fold(k), []).append(s)
print('[cache] %d uploads by %d credited artists' % (len(cache), len(by_artist)))

found = {}
for t in tracks:
    arts = parts(t['artist_name'])
    cands = [s for a in arts for s in by_artist.get(fold(a), []) if s['sid'] and str(s['sid']).isdigit() and same(t['title'], s['title'], arts)]
    if cands:
        best = max(cands, key=lambda s: s['plays'])
        found[t['id']] = (str(best['sid']), best['url'], 'cache', best['title'])
print('[cache] linked %d' % len(found))

# ── pass 2: SoundCloud search ──────────────────────────────────────────────────
if SEARCH:
    left = [t for t in tracks if t['id'] not in found]
    env = {}
    for line in (ROOT / '.env').read_text(encoding='utf-8').splitlines():
        if '=' in line and not line.strip().startswith('#'):
            k, v = line.split('=', 1); env[k.strip()] = v.strip().strip('\'"')
    req = urllib.request.Request('https://api.supabase.com/v1/projects/%s/api-keys?reveal=true' % REF,
                                 headers={'Authorization': 'Bearer ' + env['SBP_PAT'], 'User-Agent': 'comewith-link/1.0'})
    srk = [k['api_key'] for k in json.load(urllib.request.urlopen(req)) if k.get('name') == 'service_role'][0]
    for i in range(0, len(left), 20):
        batch = left[i:i + 20]
        qs = []
        for t in batch:
            core = re.sub(r'[\(\[][^\)\]]*[\)\]]', ' ', t['title'] or '')
            qs.append({'key': t['id'], 'q': (parts(t['artist_name'])[0] if parts(t['artist_name']) else '') + ' ' + core})
        body = json.dumps({'action': 'search_tracks', 'queries': qs}).encode()
        r = urllib.request.Request('https://%s.supabase.co/functions/v1/sc-connect' % REF, data=body,
                                   headers={'Authorization': 'Bearer ' + srk, 'apikey': srk, 'Content-Type': 'application/json'})
        try:
            res = json.load(urllib.request.urlopen(r, timeout=120))
        except urllib.error.HTTPError as e:
            sys.exit('search failed: %s %s' % (e.code, e.read().decode()[:300]))
        for t in batch:
            arts = parts(t['artist_name'])
            fa = {fold(a) for a in arts}
            ok = []
            for c in res.get('results', {}).get(t['id'], []):
                up = {fold(c.get('username') or ''), fold(c.get('user_permalink') or '').replace(' ', '')}
                by_artist_upload = bool(fa & up) or any(a.replace(' ', '') in up for a in fa)
                names_artist = any(re.search(r'\b' + re.escape(a) + r'\b', fold(c.get('title') or '')) for a in fa)
                if (by_artist_upload or names_artist) and same(t['title'], c.get('title'), arts):
                    ok.append((0 if by_artist_upload else 1, -(c.get('playback_count') or 0), c))
            if ok:
                c = sorted(ok, key=lambda x: (x[0], x[1]))[0][2]
                found[t['id']] = (str(c['id']), c.get('permalink_url'), 'search', c.get('title'))
        print('[search] %d/%d asked, %d linked so far' % (min(i + 20, len(left)), len(left), len(found)))

# ── report / write ─────────────────────────────────────────────────────────────
for t in tracks:
    f = found.get(t['id'])
    print('%s SHOW %-3s %-28.28s %-40.40s %s' % ('OK  ' if f else 'MISS', t['station_no'], t['artist_name'], t['title'],
                                               ('%s [%s] %s' % (f[2], f[0], f[3]))[:70] if f else ''))
print('\n%d of %d linked (%s)' % (len(found), len(tracks), 'WRITING' if WRITE else 'dry run - pass --write'))
if WRITE and found:
    vals = ','.join("('%s'::uuid,'%s','%s','%s')" % (k, v[0], (v[1] or '').replace("'", "''"), v[2]) for k, v in found.items())
    db("""update sc_playlist_tracks t set sc_link_id = v.sid, sc_link_url = nullif(v.url, ''), sc_link_how = v.how, sc_link_checked_at = now()
          from (values %s) v(id, sid, url, how) where t.id = v.id and t.sc_link_id is null;""" % vals)
    ids = ','.join("'%s'" % t['id'] for t in tracks if t['id'] not in found)
    if ids:
        db("update sc_playlist_tracks set sc_link_checked_at = now() where id in (%s);" % ids)
    print('written.')
