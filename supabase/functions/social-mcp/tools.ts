// social-mcp tool logic - the RULES, kept apart from the transport and the
// database so they can be tested with no network:
//
//   node --test supabase/functions/social-mcp/tools.test.ts
//
// The connector is deliberately narrow. It can read posts, create a skeleton,
// write a brief and CLAUDE's caption, and add a note. It cannot delete anything,
// cannot write the final caption (`caption`), and cannot change a post that has
// reached `approved` or later. Every one of those refusals is enforced HERE, on
// top of the zod schemas in server.ts - a schema says what shape a call may
// have, these say what it may do. The final-caption guard is enforced a third
// time in the database (217's trigger makes claude_caption the connector's
// alone; the connector's own write list never names `caption`).

export const ACCOUNTS = ["come_with", "di", "collab"] as const;
export const FORMATS = ["reel", "carousel", "story", "post"] as const;
export const PHASES = ["awareness", "sponsors", "radio", "convert", "event", "post", "general"] as const;
export const STAGES = ["idea", "drafted", "ready", "approved", "scheduled", "posted"] as const;

// The only stages the connector may write to. Everything from `ready` on is a
// human's: Janelle has the copy, or Keith has signed it off.
export const DRAFTABLE = ["idea", "drafted"];

// Columns the connector is allowed to SET, ever. `caption` (the final caption)
// is not on it and must never be.
export const WRITABLE = new Set(["brief", "claude_caption", "claude_drafted_at", "stage"]);

export const NOTE_AUTHOR = "Claude";
export const OWNER_EMAIL = "janelle@comewith.org";
export const TZ = "America/New_York";
export const MAX_RANGE_DAYS = 120;

export type Post = {
  id: string; title: string; scheduled_for: string | null; posted_at?: string | null;
  account: string; format: string | null; phase: string; stage: string;
  brief: string | null; claude_caption: string | null; claude_drafted_at?: string | null;
  caption: string | null; asset_url: string | null; results: Record<string, number> | null;
  note_count?: number; deleted_at?: string | null;
};

export interface Store {
  listPosts(q: { fromIso: string; toIso: string; stage?: string; account?: string }): Promise<Post[]>;
  listPosted(q: { fromIso: string; toIso: string }): Promise<Post[]>;
  getPost(id: string): Promise<Post | null>;
  findByTitleBetween(title: string, fromIso: string, toIso: string): Promise<Post[]>;
  insertPost(row: Record<string, unknown>): Promise<Post>;
  // Conditional write: applies ONLY while the row is still at `expectStage` and
  // not deleted, so a human moving the post on mid-call wins. null = no row.
  updatePostIfStage(id: string, expectStage: string, patch: Record<string, unknown>): Promise<Post | null>;
  insertNote(postId: string, body: string, authorName: string): Promise<{ id: string; created_at: string }>;
  profileIdByEmail(email: string): Promise<string | null>;
  log(entry: { tool: string; post_id?: string | null; ok: boolean; detail?: string }): Promise<void>;
}

export class Refusal extends Error {}

// ---- dates -----------------------------------------------------------------
// "2026-10-05" means that day in New York, not in UTC - a 9pm post would
// otherwise fall into the next day's list.
function tzOffsetMinutes(at: Date): number {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: TZ, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  }).formatToParts(at);
  const g = (t: string) => Number(parts.find((p) => p.type === t)!.value);
  const asUtc = Date.UTC(g("year"), g("month") - 1, g("day"), g("hour"), g("minute"), g("second"));
  return Math.round((asUtc - at.getTime()) / 60000);
}
export function nyDayStartIso(day: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) throw new Refusal(`"${day}" is not a YYYY-MM-DD date`);
  const guess = new Date(day + "T00:00:00Z");
  if (isNaN(guess.getTime())) throw new Refusal(`"${day}" is not a real date`);
  // The offset that applies AT local midnight, not at midday: on the day the
  // clocks change the two differ (Nov 1 starts in EDT and ends in EST). Two
  // passes settle it - the first lands within an hour, the second is exact.
  let start = guess.getTime() - tzOffsetMinutes(guess) * 60000;
  start = guess.getTime() - tzOffsetMinutes(new Date(start)) * 60000;
  return new Date(start).toISOString();
}
export function nyDayOf(iso: string): string {
  return new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit" })
    .format(new Date(iso));
}
function addDays(day: string, n: number): string {
  const d = new Date(day + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
// Inclusive [from, to] in New York days -> half-open UTC instants.
export function rangeIso(from: string, to: string): { fromIso: string; toIso: string } {
  const fromIso = nyDayStartIso(from), toIso = nyDayStartIso(addDays(to, 1));
  if (toIso <= fromIso) throw new Refusal("`to` is before `from`");
  if ((Date.parse(toIso) - Date.parse(fromIso)) / 864e5 > MAX_RANGE_DAYS + 1) {
    throw new Refusal(`range is longer than ${MAX_RANGE_DAYS} days - ask for less`);
  }
  return { fromIso, toIso };
}

// ---- shaping -----------------------------------------------------------------
export function shapePost(p: Post) {
  return {
    id: p.id, title: p.title, scheduled_at: p.scheduled_for, account: p.account,
    format: p.format, phase: p.phase, stage: p.stage, brief: p.brief,
    claude_caption: p.claude_caption, final_caption: p.caption, asset_link: p.asset_url,
    results: p.results, note_count: p.note_count ?? 0,
  };
}

function assertOnly(input: Record<string, unknown>, allowed: string[], tool: string) {
  const extra = Object.keys(input).filter((k) => !allowed.includes(k));
  if (extra.length) {
    const fc = extra.find((k) => /^(caption|final_caption)$/.test(k));
    if (fc) throw new Refusal(`${tool} cannot touch the final caption - that is Janelle's. Write claude_caption instead.`);
    throw new Refusal(`${tool} does not accept: ${extra.join(", ")}`);
  }
}
function assertPatchSafe(patch: Record<string, unknown>) {
  for (const k of Object.keys(patch)) {
    if (!WRITABLE.has(k)) throw new Error(`internal: connector tried to write ${k}`);
  }
}

// ---- the tools (log_results is further down) ----------------------------------------------------------
export async function listPosts(store: Store, input: { from: string; to: string; stage?: string; account?: string }) {
  assertOnly(input, ["from", "to", "stage", "account"], "list_posts");
  const r = rangeIso(input.from, input.to);
  const rows = await store.listPosts({ ...r, stage: input.stage, account: input.account });
  return { range: { from: input.from, to: input.to, timezone: TZ }, count: rows.length, posts: rows.map(shapePost) };
}

export async function createPostSkeleton(store: Store, input: {
  title: string; scheduled_at: string; account: string; format: string; phase: string; brief?: string;
}) {
  assertOnly(input, ["title", "scheduled_at", "account", "format", "phase", "brief"], "create_post_skeleton");
  const title = input.title.trim();
  if (!title) throw new Refusal("title is empty");
  const at = new Date(input.scheduled_at);
  if (isNaN(at.getTime())) throw new Refusal("scheduled_at is not a date-time");
  const day = nyDayOf(at.toISOString());
  const r = rangeIso(day, day);
  const dupes = (await store.findByTitleBetween(title, r.fromIso, r.toIso))
    .filter((p) => p.title.trim().toLowerCase() === title.toLowerCase());
  if (dupes.length) throw new Refusal(`a post called "${title}" already exists on ${day} (id ${dupes[0].id})`);
  const owner = await store.profileIdByEmail(OWNER_EMAIL);
  const row = {
    title, scheduled_for: at.toISOString(), account: input.account, format: input.format,
    phase: input.phase, brief: input.brief?.trim() || null,
    stage: "idea", owner_id: owner,          // always idea, always Janelle's
  };
  const p = await store.insertPost(row);
  return { created: shapePost(p), owner: owner ? OWNER_EMAIL : null };
}

export async function updatePostDraft(store: Store, input: { id: string; brief?: string; claude_caption?: string }) {
  assertOnly(input, ["id", "brief", "claude_caption"], "update_post_draft");
  if (input.brief === undefined && input.claude_caption === undefined) {
    throw new Refusal("nothing to update - send a brief, a claude_caption, or both");
  }
  const p = await store.getPost(input.id);
  if (!p || p.deleted_at) throw new Refusal(`no post ${input.id}`);
  if (!DRAFTABLE.includes(p.stage)) {
    throw new Refusal(`"${p.title}" is at stage ${p.stage}; the connector only edits posts at idea or drafted`);
  }
  const patch: Record<string, unknown> = {};
  if (input.brief !== undefined) patch.brief = input.brief.trim() || null;
  if (input.claude_caption !== undefined) {
    const cap = input.claude_caption.trim();
    if (!cap) throw new Refusal("claude_caption is empty");
    patch.claude_caption = cap;
    patch.claude_drafted_at = new Date().toISOString();
    if (p.stage === "idea") patch.stage = "drafted";
  }
  assertPatchSafe(patch);
  const was = p.stage;
  const out = await store.updatePostIfStage(p.id, was, patch);
  if (!out) throw new Refusal(`"${p.title}" changed while this was being written; list it again and retry`);
  return { updated: shapePost(out), moved_to_drafted: was === "idea" && out.stage === "drafted" };
}

export async function addNote(store: Store, input: { id: string; text: string }) {
  assertOnly(input, ["id", "text"], "add_note");
  const text = input.text.trim();
  if (!text) throw new Refusal("note is empty");
  const p = await store.getPost(input.id);
  if (!p || p.deleted_at) throw new Refusal(`no post ${input.id}`);
  const n = await store.insertNote(p.id, text, NOTE_AUTHOR);
  return { note_id: n.id, post_id: p.id, author: NOTE_AUTHOR, created_at: n.created_at };
}

export async function getResults(store: Store, input: { from: string; to: string }) {
  assertOnly(input, ["from", "to"], "get_results");
  const r = rangeIso(input.from, input.to);
  const rows = await store.listPosted(r);
  return {
    range: { from: input.from, to: input.to, timezone: TZ }, count: rows.length,
    with_results: rows.filter((p) => p.results && Object.keys(p.results).length).length,
    posts: rows.map((p) => ({ ...shapePost(p), posted_at: p.posted_at ?? null })),
  };
}

// log_results writes the results field and NOTHING else, and only once a post
// is posted. Metrics given are merged over what is on file, so logging comments
// later never erases the views logged earlier.
export const RESULT_METRICS = ["views", "likes", "comments", "shares", "saves"] as const;
export const RESULTS_WRITABLE = new Set(["results"]);
export const MAX_METRIC = 1_000_000_000;

export async function logResults(store: Store, input: {
  id: string; views?: number; likes?: number; comments?: number; shares?: number; saves?: number;
}) {
  assertOnly(input, ["id", ...RESULT_METRICS], "log_results");
  const given: Record<string, number> = {};
  for (const k of RESULT_METRICS) {
    const v = (input as Record<string, unknown>)[k];
    if (v === undefined) continue;
    if (typeof v !== "number" || !Number.isInteger(v) || v < 0 || v > MAX_METRIC) {
      throw new Refusal(`${k} must be a whole number from 0 to ${MAX_METRIC}`);
    }
    given[k] = v;
  }
  if (!Object.keys(given).length) throw new Refusal(`nothing to log - send at least one of ${RESULT_METRICS.join(", ")}`);
  const p = await store.getPost(input.id);
  if (!p || p.deleted_at) throw new Refusal(`no post ${input.id}`);
  if (p.stage !== "posted") throw new Refusal(`"${p.title}" is at stage ${p.stage}; results can only be logged on a posted post`);
  const patch = { results: { ...(p.results || {}), ...given } };
  for (const k of Object.keys(patch)) {
    if (!RESULTS_WRITABLE.has(k)) throw new Error(`internal: log_results tried to write ${k}`);
  }
  const out = await store.updatePostIfStage(p.id, "posted", patch);
  if (!out) throw new Refusal(`"${p.title}" changed while this was being written; list it again and retry`);
  return { updated: shapePost(out), logged: given };
}

export const TOOLS = {
  log_results: logResults,
  list_posts: listPosts,
  create_post_skeleton: createPostSkeleton,
  update_post_draft: updatePostDraft,
  add_note: addNote,
  get_results: getResults,
} as const;
export type ToolName = keyof typeof TOOLS;

// One entry point for every call: run it, log it, and turn a refusal into an
// answer the model can read rather than a transport error.
export async function runTool(store: Store, name: string, input: Record<string, unknown>) {
  const fn = (TOOLS as Record<string, (s: Store, i: any) => Promise<unknown>>)[name];
  const postId = typeof input?.id === "string" ? input.id : null;
  if (!fn) {
    await store.log({ tool: name, post_id: postId, ok: false, detail: "unknown tool" }).catch(() => {});
    return { ok: false, error: `there is no tool called ${name}` };
  }
  try {
    const result = await fn(store, input) as Record<string, any>;
    const id = postId ?? result?.created?.id ?? null;
    await store.log({ tool: name, post_id: id, ok: true }).catch(() => {});
    return { ok: true, result };
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    await store.log({ tool: name, post_id: postId, ok: false, detail: msg.slice(0, 500) }).catch(() => {});
    if (e instanceof Refusal) return { ok: false, refused: true, error: msg };
    return { ok: false, error: "the connector hit an error: " + msg };
  }
}
