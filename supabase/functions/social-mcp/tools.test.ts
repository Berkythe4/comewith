// Rule tests for the social-mcp connector. No network, no database:
//
//   node --test supabase/functions/social-mcp/tools.test.ts
//
// Every refusal the sprint names is here: editing an approved post, touching
// the final caption, and a call to a tool that would delete. (The bad-secret
// case is transport-level and lives in server.test.ts.)
import { test } from "node:test";
import assert from "node:assert/strict";
import {
  runTool, nyDayStartIso, rangeIso, NOTE_AUTHOR, OWNER_EMAIL, WRITABLE,
  type Post, type Store,
} from "./tools.ts";

const JANELLE = "52245d29-ca9f-4ad7-8dcd-6c893459dd3b";

function fakeStore(seed: Partial<Post>[] = []) {
  let n = 0;
  const posts = new Map<string, Post>();
  const notes: { post_id: string; body: string; author_name: string }[] = [];
  const logs: any[] = [];
  const writes: Record<string, unknown>[] = [];
  const mk = (p: Partial<Post>): Post => ({
    id: p.id ?? `00000000-0000-4000-8000-${String(++n).padStart(12, "0")}`,
    title: "t", scheduled_for: null, posted_at: null, account: "come_with", format: null, phase: "general",
    stage: "idea", brief: null, claude_caption: null, caption: null, asset_url: null, results: null,
    deleted_at: null, ...p,
  } as Post);
  seed.forEach((p) => { const x = mk(p); posts.set(x.id, x); });
  const inRange = (iso: string | null | undefined, a: string, b: string) => !!iso && iso >= a && iso < b;
  const store: Store = {
    async listPosts({ fromIso, toIso, stage, account }) {
      return [...posts.values()].filter((p) => !p.deleted_at && inRange(p.scheduled_for, fromIso, toIso)
        && (!stage || p.stage === stage) && (!account || p.account === account));
    },
    async listPosted({ fromIso, toIso }) {
      return [...posts.values()].filter((p) => !p.deleted_at && p.stage === "posted"
        && inRange(p.posted_at ?? p.scheduled_for, fromIso, toIso));
    },
    async getPost(id) { return posts.get(id) ?? null; },
    async findByTitleBetween(title, a, b) {
      return [...posts.values()].filter((p) => !p.deleted_at && p.title.toLowerCase() === title.toLowerCase()
        && inRange(p.scheduled_for, a, b));
    },
    async insertPost(row) { writes.push(row); const p = mk(row as Partial<Post>); posts.set(p.id, p); return p; },
    async updatePostIfStage(id, expect, patch) {
      writes.push(patch);
      const p = posts.get(id);
      if (!p || p.deleted_at || p.stage !== expect) return null;
      Object.assign(p, patch); return p;
    },
    async insertNote(post_id, body, author_name) { notes.push({ post_id, body, author_name }); return { id: "note-" + notes.length, created_at: "now" }; },
    async profileIdByEmail(email) { return email === OWNER_EMAIL ? JANELLE : null; },
    async log(e) { logs.push(e); },
  };
  return { store, posts, notes, logs, writes };
}

const ID = (k: number) => `00000000-0000-4000-8000-${String(900 + k).padStart(12, "0")}`;

// ---- list_posts ----------------------------------------------------------------
test("list_posts: New York days, inclusive, shaped with both captions", async () => {
  const f = fakeStore([
    { id: ID(1), title: "late night", scheduled_for: "2026-10-06T01:30:00Z", caption: "final", claude_caption: "draft" }, // Oct 5, 9:30pm NY
    { id: ID(2), title: "next day", scheduled_for: "2026-10-06T14:00:00Z" },
    { id: ID(3), title: "gone", scheduled_for: "2026-10-05T14:00:00Z", deleted_at: "x" },
  ]);
  const out: any = await runTool(f.store, "list_posts", { from: "2026-10-05", to: "2026-10-05" });
  assert.equal(out.ok, true);
  assert.deepEqual(out.result.posts.map((p: any) => p.title), ["late night"]);
  const p = out.result.posts[0];
  assert.equal(p.final_caption, "final");
  assert.equal(p.claude_caption, "draft");
  assert.equal(p.scheduled_at, "2026-10-06T01:30:00Z");
  assert.ok("note_count" in p && "results" in p && "asset_link" in p);
  assert.deepEqual(f.writes, [], "a read writes nothing");
});

test("list_posts: refuses a backwards or oversized range", async () => {
  const f = fakeStore();
  assert.equal(((await runTool(f.store, "list_posts", { from: "2026-10-05", to: "2026-10-01" })) as any).refused, true);
  assert.equal(((await runTool(f.store, "list_posts", { from: "2026-01-01", to: "2026-12-31" })) as any).refused, true);
});

test("dates: New York midnight across DST", () => {
  assert.equal(nyDayStartIso("2026-10-05"), "2026-10-05T04:00:00.000Z"); // EDT
  assert.equal(nyDayStartIso("2026-12-05"), "2026-12-05T05:00:00.000Z"); // EST
  assert.deepEqual(rangeIso("2026-11-01", "2026-11-01"),
    { fromIso: "2026-11-01T04:00:00.000Z", toIso: "2026-11-02T05:00:00.000Z" }); // the 25-hour day
});

// ---- create_post_skeleton -------------------------------------------------------
const SKEL = { title: "CWR SHOW 11 release", scheduled_at: "2026-10-09T18:00:00-04:00", account: "come_with", format: "reel", phase: "radio" };

test("create_post_skeleton: always idea, always Janelle, never a caption", async () => {
  const f = fakeStore();
  const out: any = await runTool(f.store, "create_post_skeleton", { ...SKEL, brief: "Tease the drop" });
  assert.equal(out.ok, true);
  const row = f.writes[0];
  assert.equal(row.stage, "idea");
  assert.equal(row.owner_id, JANELLE);
  assert.equal(row.brief, "Tease the drop");
  assert.ok(!("caption" in row) && !("claude_caption" in row));
  assert.equal(f.logs[0].tool, "create_post_skeleton");
  assert.equal(f.logs[0].post_id, out.result.created.id, "the log names the new post");
});

test("create_post_skeleton: rejects a duplicate title on the same New York day", async () => {
  const f = fakeStore([{ id: ID(1), title: "cwr show 11 release", scheduled_for: "2026-10-09T15:00:00Z" }]);
  const out: any = await runTool(f.store, "create_post_skeleton", SKEL);
  assert.equal(out.refused, true);
  assert.match(out.error, /already exists on 2026-10-09/);
  assert.equal(f.writes.length, 0);
  // Same title, different day, is fine.
  const ok: any = await runTool(f.store, "create_post_skeleton", { ...SKEL, scheduled_at: "2026-10-10T18:00:00-04:00" });
  assert.equal(ok.ok, true);
});

test("create_post_skeleton: cannot be talked into a stage or a final caption", async () => {
  const f = fakeStore();
  for (const extra of [{ stage: "approved" }, { caption: "x" }, { owner_id: "someone" }]) {
    const out: any = await runTool(f.store, "create_post_skeleton", { ...SKEL, ...extra });
    assert.equal(out.refused, true, JSON.stringify(extra));
  }
  assert.equal(f.writes.length, 0);
});

// ---- update_post_draft -----------------------------------------------------------
test("update_post_draft: writes Claude's caption, stamps it, moves idea -> drafted", async () => {
  const f = fakeStore([{ id: ID(1), stage: "idea", caption: "Janelle's words" }]);
  const out: any = await runTool(f.store, "update_post_draft", { id: ID(1), claude_caption: "Draft copy", brief: "b" });
  assert.equal(out.ok, true);
  const p = f.posts.get(ID(1))!;
  assert.equal(p.claude_caption, "Draft copy");
  assert.ok(p.claude_drafted_at);
  assert.equal(p.stage, "drafted");
  assert.equal(p.caption, "Janelle's words", "the final caption is untouched");
  assert.equal(out.result.moved_to_drafted, true);
});

test("update_post_draft: a brief alone does not move the stage", async () => {
  const f = fakeStore([{ id: ID(1), stage: "idea" }]);
  await runTool(f.store, "update_post_draft", { id: ID(1), brief: "just a brief" });
  assert.equal(f.posts.get(ID(1))!.stage, "idea");
  assert.equal(f.posts.get(ID(1))!.claude_drafted_at, undefined);
});

test("update_post_draft: REFUSES an approved post (and everything after ready)", async () => {
  for (const stage of ["ready", "approved", "scheduled", "posted", "archived", "review", "planned"]) {
    const f = fakeStore([{ id: ID(1), stage, claude_caption: "old" }]);
    const out: any = await runTool(f.store, "update_post_draft", { id: ID(1), claude_caption: "new" });
    assert.equal(out.refused, true, stage);
    assert.match(out.error, /only edits posts at idea or drafted/);
    assert.equal(f.posts.get(ID(1))!.claude_caption, "old", stage);
    assert.equal(f.writes.length, 0, stage);
    assert.equal(f.logs[0].ok, false);
  }
});

test("update_post_draft: REFUSES the final caption, under either name", async () => {
  const f = fakeStore([{ id: ID(1), stage: "drafted", caption: "keep me" }]);
  for (const k of ["caption", "final_caption"]) {
    const out: any = await runTool(f.store, "update_post_draft", { id: ID(1), [k]: "overwrite" });
    assert.equal(out.refused, true);
    assert.match(out.error, /cannot touch the final caption/);
  }
  assert.equal(f.posts.get(ID(1))!.caption, "keep me");
  assert.equal(f.writes.length, 0);
});

test("update_post_draft: loses the race to a human, never overwrites them", async () => {
  const f = fakeStore([{ id: ID(1), stage: "drafted" }]);
  const realGet = f.store.getPost;
  f.store.getPost = async (id) => { const p = await realGet(id); f.posts.get(id)!.stage = "approved"; return p ? { ...p } : null; };
  const out: any = await runTool(f.store, "update_post_draft", { id: ID(1), claude_caption: "late" });
  assert.equal(out.refused, true);
  assert.equal(f.posts.get(ID(1))!.claude_caption, null);
});

test("update_post_draft: unknown or deleted post, and an empty call", async () => {
  const f = fakeStore([{ id: ID(2), deleted_at: "x" }]);
  assert.equal(((await runTool(f.store, "update_post_draft", { id: ID(1), brief: "x" })) as any).refused, true);
  assert.equal(((await runTool(f.store, "update_post_draft", { id: ID(2), brief: "x" })) as any).refused, true);
  assert.equal(((await runTool(f.store, "update_post_draft", { id: ID(2) })) as any).refused, true);
});

test("the connector's write list never contains the final caption", () => {
  assert.ok(!WRITABLE.has("caption"));
  assert.deepEqual([...WRITABLE].sort(), ["brief", "claude_caption", "claude_drafted_at", "stage"]);
});

// ---- add_note ----------------------------------------------------------------------
test("add_note: signed Claude, allowed on any live post (a note changes nothing)", async () => {
  const f = fakeStore([{ id: ID(1), stage: "approved" }]);
  const out: any = await runTool(f.store, "add_note", { id: ID(1), text: "Planned 3 posts for next week." });
  assert.equal(out.ok, true);
  assert.deepEqual(f.notes, [{ post_id: ID(1), body: "Planned 3 posts for next week.", author_name: NOTE_AUTHOR }]);
  assert.equal(f.posts.get(ID(1))!.stage, "approved");
});

// ---- get_results --------------------------------------------------------------------
test("get_results: posted only, by posted date, with results", async () => {
  const f = fakeStore([
    { id: ID(1), stage: "posted", posted_at: "2026-09-28T16:00:00Z", results: { views: 1200, likes: 80 } },
    { id: ID(2), stage: "posted", posted_at: "2026-09-20T16:00:00Z", results: { views: 5 } },
    { id: ID(3), stage: "approved", scheduled_for: "2026-09-28T16:00:00Z" },
    { id: ID(4), stage: "posted", posted_at: null, scheduled_for: "2026-09-27T16:00:00Z" },
  ]);
  const out: any = await runTool(f.store, "get_results", { from: "2026-09-26", to: "2026-10-02" });
  assert.deepEqual(out.result.posts.map((p: any) => p.id).sort(), [ID(1), ID(4)]);
  assert.equal(out.result.with_results, 1);
  assert.deepEqual(out.result.posts.find((p: any) => p.id === ID(1)).results, { views: 1200, likes: 80 });
});

// ---- no deletes ------------------------------------------------------------------------
test("there is no delete: any delete-shaped call is refused and logged", async () => {
  const f = fakeStore([{ id: ID(1) }]);
  for (const name of ["delete_post", "remove_post", "delete_note", "update_post"]) {
    const out: any = await runTool(f.store, name, { id: ID(1) });
    assert.equal(out.ok, false);
    assert.match(out.error, /no tool called/);
  }
  assert.ok(f.posts.has(ID(1)));
  assert.equal(f.logs.length, 4);
  assert.ok(f.logs.every((l) => l.ok === false));
});
