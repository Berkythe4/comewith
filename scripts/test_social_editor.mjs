// Social Calendar v2 (217): the post editor and list, tested against the real
// code lifted out of dashboard.html - no browser, no network.
//
//   node scripts/test_social_editor.mjs        (from the repo root)
//
// What it holds the editor to:
//   * every option list matches the CHECK constraint 217 put on prod
//   * a save writes ONLY the fields the editor shows - hidden fields (channels,
//     series, event, owner, CTA, ...) are never in the patch, so their data stays
//   * the editor never writes claude_caption (the connector's alone)
//   * results are written only at stage posted
//   * Approve = stage approved + a note + the bell notification to the owner
//   * the 🤖 rule, and legacy stages survive a dropdown
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];
let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);
const ok = (cond, m) => (cond ? pass(m) : fail(m));
const slice = (a, b) => {
  const i = mod.indexOf(a); const j = mod.indexOf(b, i + 1);
  if (i < 0 || j < 0) throw new Error('marker not found: ' + (i < 0 ? a : b));
  return mod.slice(i, j);
};

// ---- constants, evaluated for real -------------------------------------------
const consts = new Function(slice('const SOCIAL_STAGES = ', 'const SOCIAL_CHANNELS') +
  '\nreturn { SOCIAL_STAGES, SOCIAL_LEGACY_STAGES, SOCIAL_STAGE_LABEL, socialStageList, socialStagesPresent, SOCIAL_ACCOUNTS, SOCIAL_FORMATS, SOCIAL_PHASES, SOCIAL_RESULT_KEYS, socialNeedsReview };')();

// Transcribed from 217_social_calendar_v2.sql (applied to prod 2026-10-02).
const CHECK = {
  account: ['come_with', 'di', 'collab'],
  format: ['reel', 'carousel', 'story', 'post'],
  phase: ['awareness', 'sponsors', 'radio', 'convert', 'event', 'post', 'general'],
  stage: ['idea', 'drafted', 'ready', 'approved', 'scheduled', 'posted', 'review', 'planned', 'archived'],
};
const same = (a, b) => a.length === b.length && a.every(x => b.includes(x));
ok(same(consts.SOCIAL_ACCOUNTS, CHECK.account), 'Account options match the CHECK constraint');
ok(same(consts.SOCIAL_FORMATS, CHECK.format), 'Format options match the CHECK constraint');
ok(same(consts.SOCIAL_PHASES, CHECK.phase), 'Phase options match the CHECK constraint');
ok(JSON.stringify(consts.SOCIAL_STAGES) === JSON.stringify(['idea', 'drafted', 'ready', 'approved', 'scheduled', 'posted']),
  'Stage pipeline is idea > drafted > ready > approved > scheduled > posted');
ok(consts.SOCIAL_STAGES.concat(consts.SOCIAL_LEGACY_STAGES).every(s => CHECK.stage.includes(s) && consts.SOCIAL_STAGE_LABEL[s]),
  'every stage offered or kept is legal and labelled');
ok(consts.socialStageList('planned').includes('planned') && !consts.socialStageList('idea').includes('planned'),
  'a legacy stage stays in its own row\'s dropdown (touching it cannot rewrite it), and is not offered elsewhere');
ok(consts.socialStagesPresent([{ stage: 'archived' }]).includes('archived') && !consts.socialStagesPresent([]).includes('archived'),
  'board shows a legacy column only while a post still holds it');

// ---- the robot rule --------------------------------------------------------------
const nr = consts.socialNeedsReview;
ok(nr({ claude_caption: 'x', caption: null }) && nr({ claude_caption: 'x', caption: '   ' }), '🤖 shows when Claude drafted and Final is empty');
ok(!nr({ claude_caption: 'x', caption: 'final' }) && !nr({ claude_caption: null, caption: null }), '🤖 hides once Final is written, or when Claude has not drafted');

// ---- savePost, run against fakes -------------------------------------------------
const saveSrc = slice('async function savePost(', 'async function loadPostNotes(');
function harness(form, opts = {}) {
  const calls = { update: null, insert: null, eqId: null, notes: [], notify: [], closed: false, toast: [] };
  const sb = {
    from: (t) => ({
      update: (patch) => { calls.update = { t, patch }; return { eq: async (_c, id) => { calls.eqId = id; return { error: null }; } }; },
      insert: async (patch) => { calls.insert = { t, patch }; return { error: null }; },
    }),
  };
  const kpiVal = (n) => (form[n] == null ? '' : String(form[n]).trim());
  const fn = new Function('sb', 'kpiVal', 'SOCIAL_RESULT_KEYS', 'socialPeople', 'ME', 'notify', 'addPostNoteBody', 'closeKpi', 'toast', 'document', 'loadSocialCalendar',
    saveSrc + '\nreturn savePost;')(
    sb, kpiVal, consts.SOCIAL_RESULT_KEYS,
    [{ id: 'keith', email: 'berky@comewith.org' }, { id: 'janelle', email: 'janelle@comewith.org' }],
    { id: opts.me || 'keith' },
    async (...a) => { calls.notify.push(a); },
    async (id, body) => { calls.notes.push([id, body]); },
    () => { calls.closed = true; },
    (m) => calls.toast.push(m),
    { querySelector: () => null },
    async () => {},
  );
  return { savePost: fn, calls };
}
const FORM = {
  title: 'CWR SHOW 11 release', account: 'di', format: 'reel', scheduled_for: '2026-10-09T18:00', phase: 'radio',
  brief: 'tease it', caption: 'Final words', asset_url: 'https://drive.google.com/x', stage: 'drafted',
  res_views: '1200', res_likes: '', res_shares: '7', res_saves: 'abc',
};
const VISIBLE = ['title', 'stage', 'account', 'format', 'phase', 'brief', 'caption', 'asset_url', 'scheduled_for'];
const HIDDEN = ['channels', 'series', 'event_id', 'owner_id', 'draft_by', 'review_by', 'content_pillar', 'cta', 'link_url', 'asset_status', 'station_id'];
const EXISTING = { id: 'p1', owner_id: 'janelle', posted_at: null, channels: ['instagram'], series: 'Come With Radio' };

{
  const h = harness(FORM);
  await h.savePost(EXISTING, null, null, EXISTING);
  const keys = Object.keys(h.calls.update.patch);
  ok(same(keys, VISIBLE), 'an edit writes exactly the visible fields: ' + keys.join(', '));
  ok(!HIDDEN.some(k => k in h.calls.update.patch), 'hidden fields are never in the patch - channels/series/event/owner/CTA data stays on the row');
  ok(!('claude_caption' in h.calls.update.patch) && !('claude_drafted_at' in h.calls.update.patch), 'the editor never writes claude_caption');
  ok(!('results' in h.calls.update.patch), 'results are not written before stage posted');
  ok(h.calls.update.patch.caption === 'Final words' && h.calls.update.patch.account === 'di', 'Final caption is the existing caption column');
  ok(h.calls.notify.length === 0, 'an ordinary save notifies nobody');
}
{
  const h = harness({ ...FORM, stage: 'posted' });
  await h.savePost(EXISTING, null, null, EXISTING);
  ok(JSON.stringify(h.calls.update.patch.results) === JSON.stringify({ views: 1200, shares: 7 }),
    'at posted, results keep the numbers entered and drop blanks/junk: ' + JSON.stringify(h.calls.update.patch.results));
  ok(!!h.calls.update.patch.posted_at, 'marking posted stamps posted_at');
}
{
  const h = harness({ ...FORM, stage: 'posted', res_comments: '12' });
  await h.savePost({ ...EXISTING, results: { views: 1, comments: 3, reach: 5000 } }, null, null, EXISTING);
  const r = h.calls.update.patch.results;
  ok(r.comments === 12 && r.views === 1200, 'Comments is an editor box like the other four');
  ok(r.reach === 5000, 'a metric the editor has no box for survives a save (never wiped)');
  ok(JSON.stringify(consts.SOCIAL_RESULT_KEYS) === JSON.stringify(['views', 'likes', 'comments', 'shares', 'saves']),
    'editor result boxes = the connector log_results metrics');
}
{
  const h = harness({ ...FORM, stage: 'posted', res_views: '', res_shares: '', res_saves: '' });
  await h.savePost({ ...EXISTING, posted_at: '2026-10-01T00:00:00Z' }, null, null, EXISTING);
  ok(h.calls.update.patch.results === null && !('posted_at' in h.calls.update.patch), 'all results cleared -> null; an existing posted_at is kept');
}
{
  const h = harness(FORM);
  await h.savePost(EXISTING, null, { stage: 'approved', copy_status: 'approved', approve: true }, EXISTING);
  ok(h.calls.update.patch.stage === 'approved' && h.calls.update.patch.copy_status === 'approved', 'Approve saves the form at stage approved');
  ok(h.calls.notify.length === 1 && h.calls.notify[0][0] === 'janelle' && h.calls.notify[0][1] === 'approved' && h.calls.notify[0][4] === 'post',
    'Approve fires the bell to the post owner (kind approved, subject post)');
  ok(h.calls.notes.length === 1 && /Approved/.test(h.calls.notes[0][1]), 'Approve leaves a note in the thread');
  ok(h.calls.closed, 'the modal closes after the save');
}
{
  const h = harness(FORM, { me: 'janelle' });
  await h.savePost(EXISTING, null, { stage: 'approved', approve: true }, EXISTING);
  ok(h.calls.notify.length === 0, 'approving your own post does not notify yourself');
}
{
  const h = harness({ ...FORM, format: '' });
  await h.savePost(null, null, null, { event_id: 'ev1', series: 'Dance Infusion', channels: ['instagram'] });
  const p = h.calls.insert.patch;
  ok(p.owner_id === 'janelle', 'a new post is Janelle\'s by default (so Approve has somebody to notify)');
  ok(p.event_id === 'ev1' && p.series === 'Dance Infusion' && p.channels[0] === 'instagram', 'a new post keeps what its seed carried (event hub / content center)');
  ok(p.format === null, 'an unpicked format saves as null, never a guess');
}
{
  const h = harness({ ...FORM, title: '' });
  let threw = false; try { await h.savePost(EXISTING, null, null, EXISTING); } catch { threw = true; }
  ok(threw && !h.calls.update, 'no title, no save');
}

// ---- the editor template ------------------------------------------------------------
const modal = slice('async function openPostModal(post, opts) {', 'function spFormSnapshot()');
for (const f of ['title', 'account', 'format', 'scheduled_for', 'phase', 'brief', 'caption', 'asset_url', 'stage']) {
  if (!modal.includes(`name="${f}"`) && !modal.includes(`spSeg('${f}'`)) fail('editor is missing ' + f);
}
pass('editor shows title, account, format, date/time, phase, brief, final caption, asset link, stage');
ok(!/name="(channels|series|event_id|owner_id|draft_by|review_by|cta|link_url|asset_status|content_pillar)"|data-chan|data-post-req|postTaskList|postAssetList/.test(modal),
  'hidden: channels, series, event, owner, drafter/approver, CTA, destination URL, asset status, tasks, attached content, workflow buttons');
ok(modal.includes('Claude drafts this on Mondays for posts in the next 14 days.') || mod.includes("SP_CLAUDE_EMPTY = 'Claude drafts this on Mondays for posts in the next 14 days.'"),
  'empty Claude caption explains when it arrives');
ok(/data-sp-use/.test(modal) && /data-sp-copy="claude"/.test(modal) && /data-sp-copy="final"/.test(modal), 'Use this / Copy / Copy final caption buttons present');
ok(/<details class="sp-notes/.test(modal) && !/<details[^>]*\bopen\b/.test(modal) && /spNoteCount/.test(modal), 'notes are collapsed by default and show a count');
ok(/id="spResults" style="\$\{stage === 'posted' \? '' : 'display:none;'\}"/.test(modal), 'results only visible at stage posted');
ok(/kpiCloseGuard = \(\) => spFormSnapshot\(\) !== snap/.test(modal), 'unsaved-changes guard is armed when the editor opens');
const clicks = slice("const cp = e.target.closest('#kpiModalBody [data-sp-copy]');", "const use = e.target.closest");
ok(/kpiVal\('caption'\) \|\| claudeTxt/.test(clicks), 'Copy final caption falls back to Claude\'s caption when Final is empty');
ok(/setTimeout\(\(\) => \{ btn\.textContent = btn\.dataset\.label;[^}]*\}, 2000\)/.test(mod), 'Copied label reverts after 2 seconds');

// ---- the close guard -----------------------------------------------------------------
ok(/\$\('kpiModalX'\)\.addEventListener\('click', requestCloseKpi\)/.test(mod) &&
   /e\.key === 'Escape' && \$\('kpiModalOverlay'\)\.classList\.contains\('show'\)\) requestCloseKpi\(\)/.test(mod),
  'Esc and ✕ go through the guard');
ok(!/function requestCloseKpi[\s\S]{0,900}confirm\(/.test(mod), 'the guard is in the modal, not a browser confirm()');

// ---- the list -------------------------------------------------------------------------
const list = slice('function socialListHTML(posts) {', 'function renderSocialFilters()');
ok(!/socialChanCell|content_pillar/.test(list) && /socialTagChips\(p\)/.test(list), 'list: Channels and Pillar columns replaced by Account/Format/Phase chips');
ok(/socialNeedsReview\(p\)/.test(list) && list.toLowerCase().includes('\\ud83e\\udd16'),'list: 🤖 on posts that need Janelle');
ok(/data-sp-field="stage"/.test(list), 'list: stage dropdown still saves on change');
const filters = slice('function renderSocialFilters()', 'async function socialPatch(');
ok(/group\('Phase'/.test(filters) && /group\('Account'/.test(filters) && /group\('Stage'/.test(filters), 'filters: Phase, Account and Stage chips');
const cols = (list.match(/<th[ >]/g) || []).length, span = (list.match(/colspan="(\d+)"/) || [])[1];
ok(String(cols) === span, `empty state spans all ${cols} columns`);

console.log(fails ? `\n${fails} FAILED` : '\nALL PASS');
process.exit(fails ? 1 : 0);
