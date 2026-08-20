// The staged/live rule for recap videos exists in TWO places — v_public_recap
// (migration 184) and the dashboard — and they have to agree exactly, or a video
// you staged in the editor still renders on the site. This asserts the JS half
// against the SQL half, which is transcribed below.
//
//   SQL:  where coalesce((v ->> 'is_public')::boolean, true)
//   i.e.  a MISSING flag means PUBLIC — that is what keeps every existing video
//         rendering with no backfill.
//
//   node scripts/test_recap_publish.mjs        (from the repo root)
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];
const sql = fs.readFileSync('supabase/migrations/184_recap_video_publish.sql', 'utf8');

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);

// ---- the two halves must state the same rule --------------------------------
if (!/coalesce\(\(t\.v ->> 'is_public'\)::boolean, true\)/.test(sql)) {
  fail('the view no longer defaults a missing flag to public');
} else pass('SQL: a missing is_public means public');

const m = mod.match(/const recapIsPublic = [^;]+;/);
if (!m) fail('recapIsPublic() is not defined');
else {
  const recapIsPublic = new Function('return ' + m[0].replace('const recapIsPublic = ', '').replace(/;$/, ''))();
  const cases = [
    [{ url: 'x' }, true, 'no flag at all is public — every existing video'],
    [{ url: 'x', is_public: true }, true, 'explicit true'],
    [{ url: 'x', is_public: false }, false, 'explicit false is staged'],
    [{ url: 'x', is_public: undefined }, true, 'undefined is public'],
    [{ url: 'x', is_public: null }, true, 'null is public, matching coalesce'],
  ];
  let bad = 0;
  for (const [v, want, why] of cases) {
    if (recapIsPublic(v) !== want) { fail('recapIsPublic(' + JSON.stringify(v) + ') should be ' + want + ' — ' + why); bad++; }
  }
  if (!bad) pass('JS: same rule, all ' + cases.length + ' cases');
}

// ---- the thumbnail rule ------------------------------------------------------
// events.youtube_url drives the homepage card image. A staged video must not
// leave a broken one, so both halves pick the first PUBLIC YouTube link.
if (!/and t\.v ->> 'url' ~\* '\(youtube\\\.com\|youtu\\\.be\)'/.test(sql)) {
  fail('the view does not restrict the thumbnail to YouTube links');
} else pass('SQL: thumbnail comes from a public YouTube entry');

const jsThumb = mod.match(/patch\.youtube_url = \(recapVideos\.find\([^)]*\)[^;]*;/);
if (!jsThumb || !/v\.is_public !== false/.test(jsThumb[0])) {
  fail('the editor still takes the first YouTube link regardless of whether it is staged');
} else pass('editor: thumbnail skips staged videos');

if (!/recapIsPublic\(v\) && \/youtu\\\.\?be\/i\.test/.test(mod)) {
  fail('the Content tab writer does not skip staged videos for the thumbnail');
} else pass('Content tab: thumbnail skips staged videos');

// ---- the editor must not repeat the promise it used to break -----------------
if (/stay hidden on the site until fixed/.test(mod)) {
  fail('the old confirm still claims a private link stays hidden — it did not');
} else pass('the untrue "stored but hidden" confirm is gone');
if (!/v\.is_public = false;/.test(mod)) fail('an unembeddable link is not forced to staged on save');
else pass('a link that will not embed is saved staged, which makes the promise true');

// ---- new rows start staged ---------------------------------------------------
if (!/recapVidRow\('Watch the recap', '', '', false\)/.test(mod)) {
  fail('a newly added editor row does not start staged');
} else pass('new editor rows start staged');

// ---- the Content tab is editable in place -----------------------------------
for (const need of ['data-cv-pub', 'data-cv-label', 'data-cv-del', 'data-cv-add']) {
  if (!mod.includes(need)) fail('the Content tab is missing ' + need);
}
if (!fails) pass('Content tab: label, live/staged, remove and add are all in place');

// content_assets is admin-only — nothing public reads it, so it must not claim
// anything is "on site".
if (/a\.is_public \? ' · 🌐 on site' : ''/.test(mod)) {
  fail('the clip library still claims rows are on the site; nothing public reads content_assets');
} else pass('the clip library no longer claims to be on the site');

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
