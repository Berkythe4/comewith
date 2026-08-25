// "Email task list" on Calendar & Tasks sends ONLY what the board is currently
// showing — the active filters, in the order on screen. That is a correctness
// claim, not a cosmetic one: a filtered list emailed as though it were the whole
// list is how someone concludes there is no outstanding work.
//
// This pulls the real functions out of dashboard.html and runs them, rather than
// grepping for their source, so a refactor that changes behaviour fails here.
//
//   node scripts/test_task_email.mjs        (from the repo root)
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

const src = fs.readFileSync('dashboard.html', 'utf8').replace(/\r\n/g, '\n');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);
const eq = (got, want, m) => (JSON.stringify(got) === JSON.stringify(want)
  ? pass(m) : fail(`${m}\n        got  ${JSON.stringify(got)}\n        want ${JSON.stringify(want)}`));

// ---- lift the functions under test, with the browser bits stubbed ----------
const grab = (startMark, endMark) => {
  const i = mod.indexOf(startMark);
  if (i < 0) throw new Error('could not find ' + startMark);
  const j = mod.indexOf(endMark, i);
  if (j < 0) throw new Error('could not find end ' + endMark);
  return mod.slice(i, j);
};
const builder = grab('function buildTasksEmailHtml(tasks, opts = {})', '\nasync function hubEmailTasks(');
const calBits = grab('const CAL_DUE_WORDS', '\nasync function calEmailTasks(');
const board = grab('function calBoardTasks() {', '\nfunction calRenderBoard(');

const STUBS = `
const CAL_DEFAULT_STATUS = ['todo','doing','blocked'];
const CAL_PRIO_RANK = { high: 0, medium: 1, low: 2 };
const wsPillarLabel = (p) => ({ dance_infusion: 'Dance Infusion', audience: 'Audience', ops: 'Operations' }[p] || p);
const escapeHtml = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmtDate = (d) => new Date(d + 'T00:00:00').toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' });
const calToday = () => TODAY;
const calAddDays = (ds, n) => { const d = new Date(ds+'T00:00:00'); d.setDate(d.getDate()+n); return d.toISOString().slice(0,10); };
export const CAL = { tasks: [], board: null };
export const TODAY = '2026-08-25';
`;
const file = path.join(os.tmpdir(), 'cw_task_email_' + process.pid + '.mjs');
fs.writeFileSync(file, STUBS + builder + '\n' + calBits + '\n' + board +
  '\nexport { buildTasksEmailHtml, calFilterWords, calBoardTasks };\n');
const M = await import('file://' + file.replace(/\\/g, '/'));
fs.unlinkSync(file);

// ---- fixture ---------------------------------------------------------------
const A = (id, name) => ({ id: 'as' + id, role: 'doer', actor: { id, display_name: name } });
const TASKS = [
  { id: 't1', title: 'Confirm Ali on the lineup', description: 'Waiting on a yes', status: 'doing', priority: 'high',
    due_date: '2026-08-20', milestone: false, pillar: 'dance_infusion', event: { id: 'e1', name: 'Dance Infusion 3' }, task_assignments: [A('m1', 'Martin')] },
  { id: 't2', title: 'Post the flyer', status: 'todo', priority: 'medium', due_date: '2026-08-28',
    milestone: true, pillar: 'audience', event: { id: 'e1', name: 'Dance Infusion 3' }, task_assignments: [A('l1', 'Liz'), A('m1', 'Martin')] },
  { id: 't3', title: 'Book the sound tech', status: 'todo', priority: null, due_date: null,
    milestone: false, pillar: null, event: { id: 'e2', name: 'Come With 9-12' }, task_assignments: [] },
  { id: 't4', title: 'Pay the venue deposit', status: 'done', priority: 'high', due_date: '2026-08-10',
    milestone: false, pillar: 'ops', event: { id: 'e1', name: 'Dance Infusion 3' }, task_assignments: [A('k1', 'Keith')] },
  { id: 't5', title: 'Chase the photographer', status: 'blocked', priority: 'low', due_date: '2026-09-30',
    milestone: false, pillar: null, event: { id: 'e2', name: 'Come With 9-12' }, task_assignments: [A('k1', 'Keith')] },
  // Deliberately 'doing' but NOT overdue: grouped mode files overdue tasks under
  // Overdue whatever their status, so without this there is no In-progress section.
  { id: 't6', title: 'Draft the run of show', status: 'doing', priority: 'medium', due_date: '2026-09-05',
    milestone: false, pillar: 'ops', event: { id: 'e2', name: 'Come With 9-12' }, task_assignments: [A('l1', 'Liz')] },
];
const base = () => ({ q: '', status: ['todo', 'doing', 'blocked'], event: '', assignee: '', priority: '',
                      pillar: '', due: 'all', msOnly: false, unassigned: false, sort: 'due', dir: 1 });
const withBoard = (mut) => { M.CAL.tasks = TASKS; M.CAL.board = base(); mut(M.CAL.board); return M.CAL.board; };
// A milestone renders as "⭐ Title", so match the title with an optional star.
const at = (html, t) => { const i = html.indexOf('>' + t.title + '<'); return i >= 0 ? i : html.indexOf('>⭐ ' + t.title + '<'); };
const titlesIn = (html) => TASKS.filter(t => at(html, t) >= 0).map(t => t.title);
const orderIn = (html) => TASKS.map(t => [t.title, at(html, t)])
  .filter(([, i]) => i >= 0).sort((a, b) => a[1] - b[1]).map(([n]) => n);

// ---- 1. the emailed set IS the filtered set, never more ---------------------
{
  withBoard(f => { f.priority = 'high'; });
  const shown = M.calBoardTasks();
  const html = M.buildTasksEmailHtml(shown, { order: 'as-sorted', showEvent: true });
  eq(titlesIn(html).sort(), shown.map(t => t.title).sort(), 'the email contains exactly the filtered rows (priority: high)');
  if (html.includes('Pay the venue deposit')) fail('a task the filter excluded reached the email');
  else pass('a done task outside the filter is not in the email');
}
{
  withBoard(f => { f.assignee = 'k1'; f.status = ['todo', 'doing', 'blocked', 'done']; });
  const shown = M.calBoardTasks();
  const html = M.buildTasksEmailHtml(shown, { order: 'as-sorted' });
  eq(titlesIn(html).sort(), ['Chase the photographer', 'Pay the venue deposit'], 'assignee filter carries through to the email');
}
{
  withBoard(f => { f.q = 'flyer'; });
  const html = M.buildTasksEmailHtml(M.calBoardTasks(), { order: 'as-sorted' });
  eq(titlesIn(html), ['Post the flyer'], 'the search box narrows the email too');
}

// ---- 2. as-sorted keeps the ON-SCREEN order, unchanged ---------------------
{
  const f = withBoard(() => {});
  f.sort = 'due'; f.dir = 1;
  const asc = M.calBoardTasks();
  const htmlAsc = M.buildTasksEmailHtml(asc, { order: 'as-sorted' });
  eq(orderIn(htmlAsc), asc.map(t => t.title), 'due ascending: email order matches the board exactly');

  f.dir = -1;
  const desc = M.calBoardTasks();
  const htmlDesc = M.buildTasksEmailHtml(desc, { order: 'as-sorted' });
  eq(orderIn(htmlDesc), desc.map(t => t.title), 'due descending: email order matches the board exactly');
  if (JSON.stringify(orderIn(htmlAsc)) === JSON.stringify(orderIn(htmlDesc))) fail('reversing the sort did not change the email');
  else pass('reversing the board sort reverses the email');

  f.sort = 'priority'; f.dir = 1;
  const byPrio = M.calBoardTasks();
  eq(orderIn(M.buildTasksEmailHtml(byPrio, { order: 'as-sorted' })), byPrio.map(t => t.title), 'priority sort carries through to the email');
}

// ---- 3. grouped mode regroups; as-sorted must not --------------------------
{
  const all = TASKS.slice();
  const grouped = M.buildTasksEmailHtml(all, { order: 'grouped', includeDone: true });
  for (const label of ['Overdue', 'In progress', 'Blocked', 'To do']) {
    if (!grouped.includes(label)) fail('grouped mode is missing the ' + label + ' section');
  }
  pass('grouped mode still sections by status');
  const flat = M.buildTasksEmailHtml(all, { order: 'as-sorted' });
  if (flat.includes('In progress') || flat.includes('To do')) fail('as-sorted mode grouped the list anyway');
  else pass('as-sorted mode emits one flat list, no status sections');
  eq(orderIn(flat), all.map(t => t.title), 'as-sorted preserves the exact array order it was handed');
}

// ---- 4. a filtered email SAYS it is filtered --------------------------------
{
  withBoard(f => { f.priority = 'high'; f.due = 'overdue'; });
  const words = M.calFilterWords();
  if (!words.length) fail('calFilterWords reported no active filters when two are set');
  const note = `This is a filtered view of the task board (1 of 5 tasks), not the full list — ${words.join(' · ')}.`;
  const html = M.buildTasksEmailHtml(M.calBoardTasks(), { order: 'as-sorted', filterNote: note });
  if (!html.includes('filtered view of the task board')) fail('the filter note is not in the email body');
  else pass('a filtered email states that it is filtered');

  withBoard(() => {});
  eq(M.calFilterWords(), [], 'an unfiltered board reports no filter words');
  const plain = M.buildTasksEmailHtml(M.calBoardTasks(), { order: 'as-sorted', filterNote: '' });
  if (plain.includes('filtered view of the task board')) fail('an unfiltered email claims to be filtered');
  else pass('an unfiltered email carries no filter note');
}

// ---- 5. the filter words actually name each filter --------------------------
{
  const cases = [
    [f => { f.priority = 'high'; }, 'high priority'],
    [f => { f.pillar = 'dance_infusion'; }, 'bucket: Dance Infusion'],
    [f => { f.due = 'overdue'; }, 'overdue only'],
    [f => { f.due = '7'; }, 'overdue + next 7 days'],
    [f => { f.msOnly = true; }, 'milestones only'],
    [f => { f.unassigned = true; }, 'unassigned only'],
    [f => { f.q = 'flyer'; }, 'matching “flyer”'],
    [f => { f.assignee = 'm1'; }, 'assigned to Martin'],
    [f => { f.event = 'e1'; }, 'event: Dance Infusion 3'],
    [f => { f.status = ['done']; }, 'status: done'],
  ];
  let bad = 0;
  for (const [mut, want] of cases) {
    withBoard(mut);
    if (!M.calFilterWords().includes(want)) { bad++; fail(`calFilterWords should say "${want}" — got ${JSON.stringify(M.calFilterWords())}`); }
  }
  if (!bad) pass(`all ${cases.length} filters are named in words for the email`);
}

// ---- 6. cross-event lists must say which event each task belongs to ---------
{
  const html = M.buildTasksEmailHtml(TASKS.slice(0, 3), { order: 'as-sorted', showEvent: true });
  if (!html.includes('Dance Infusion 3') || !html.includes('Come With 9-12')) fail('showEvent did not put event names on the rows');
  else pass('a cross-event list names each task\'s event');
  const hubStyle = M.buildTasksEmailHtml(TASKS.slice(0, 3), { order: 'grouped' });
  if (hubStyle.includes('Dance Infusion 3')) fail('the single-event email repeats the event name on every row');
  else pass('the single-event (hub) email leaves the event off the rows');
}

// ---- 7. assignee and due date are on every row ------------------------------
{
  const html = M.buildTasksEmailHtml(TASKS.slice(0, 3), { order: 'as-sorted', showEvent: true });
  if (!html.includes('Martin')) fail('the assignee is missing from the email row');
  else if (!html.includes('Liz, Martin')) fail('multiple assignees are not both listed');
  else if (!html.includes('unassigned')) fail('an unassigned task does not say so');
  else pass('every row carries its assignees (or says unassigned)');
  if (!html.includes('due Aug 20')) fail('the due date is missing from the email row');
  else pass('every dated row carries its due date');
  if (!/color:#C75A3C;font-weight:bold;">due Aug 20/.test(html)) fail('an overdue task is not flagged in the email');
  else pass('an overdue task is flagged red in the email');
}

// ---- 8. header counts describe the list that was sent ----------------------
{
  const html = M.buildTasksEmailHtml(TASKS, { order: 'as-sorted' });
  if (!/<b>5<\/b> open/.test(html)) fail('the header open-count does not match the list');
  else if (!/<b style="color:#C75A3C;">1 overdue<\/b>/.test(html)) fail('the header overdue-count does not match the list');
  else if (!/1 done/.test(html)) fail('the header done-count does not match the list');
  else pass('the header counts describe the tasks actually sent');
}

console.log(fails ? `\n${fails} FAILED` : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
