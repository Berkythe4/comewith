// Transport tests for social-mcp: the real MCP handler, driven with the same
// HTTP requests claude.ai sends, over an in-memory store. No network.
//
//   cd supabase/functions/social-mcp/.test-deps && npm i @modelcontextprotocol/server@2.2.0 zod@4
//   node --import ./supabase/functions/social-mcp/npm-loader.mjs --test supabase/functions/social-mcp/server.test.ts
import { test } from "node:test";
import assert from "node:assert/strict";
import { makeFetch, pathSecret, secretMatches } from "./server.ts";
import type { Post, Store } from "./tools.ts";

const SECRET = "s".repeat(20) + "0123456789abcdefABCDEF";
const BASE = "https://x.supabase.co/functions/v1/social-mcp";
const PID = "00000000-0000-4000-8000-000000000001";

const LOGS: any[] = [];
function store(): Store & { posts: Map<string, Post> } {
  const posts = new Map<string, Post>([[PID, {
    id: PID, title: "Approved thing", scheduled_for: "2026-10-06T16:00:00Z", account: "come_with", format: "reel",
    phase: "radio", stage: "approved", brief: null, claude_caption: null, caption: "final", asset_url: null, results: null,
  } as Post]]);
  return {
    posts,
    async listPosts() { return [...posts.values()]; },
    async listPosted() { return []; },
    async getPost(id) { return posts.get(id) ?? null; },
    async findByTitleBetween() { return []; },
    async insertPost(r) { const p = { ...(r as any), id: "00000000-0000-4000-8000-000000000002" }; posts.set(p.id, p); return p; },
    async updatePostIfStage(id, s, patch) { const p = posts.get(id); if (!p || p.stage !== s) return null; Object.assign(p, patch); return p; },
    async insertNote() { return { id: "n1", created_at: "now" }; },
    async profileIdByEmail() { return "janelle"; },
    async log(e) { LOGS.push(e); },
  };
}

const s = store();
const handle = makeFetch(() => s, SECRET);
let rpcId = 0;
async function rpc(url: string, method: string, params: unknown = {}, headers: Record<string, string> = {}) {
  const res = await handle(new Request(url, {
    method: "POST",
    headers: {
      "content-type": "application/json", accept: "application/json, text/event-stream",
      "mcp-protocol-version": "2025-06-18", ...headers,
    },
    body: JSON.stringify({ jsonrpc: "2.0", id: ++rpcId, method, params }),
  }));
  const text = await res.text();
  let body: any = null;
  try { body = JSON.parse(text); } catch {
    const line = text.split("\n").find((l) => l.startsWith("data:"));
    body = line ? JSON.parse(line.slice(5)) : text;
  }
  return { status: res.status, body };
}
const callTool = (name: string, args: unknown, url = `${BASE}/${SECRET}`) =>
  rpc(url, "tools/call", { name, arguments: args });

test("bad, missing or short secret: bare 404, nothing reached", async () => {
  for (const url of [`${BASE}/wrong-secret-wrong-secret-wrong-secret-x`, BASE, `${BASE}/`, `${BASE}/${SECRET.slice(0, -1)}`]) {
    const r = await rpc(url, "tools/list");
    assert.equal(r.status, 404, url);
    assert.equal(r.body, "Not found");
  }
  assert.equal(secretMatches("short", "short"), false, "a short secret is never accepted, even if it matches");
});

test("secret in the path or in an x-api-key header", async () => {
  assert.equal(pathSecret(`${BASE}/${SECRET}/mcp`), SECRET);
  assert.equal(pathSecret(`https://h/social-mcp/${SECRET}`), SECRET);
  assert.equal((await rpc(`${BASE}/${SECRET}`, "tools/list")).status, 200);
  assert.equal((await rpc(`${BASE}/${SECRET}/mcp`, "tools/list")).status, 200);
  assert.equal((await rpc(BASE, "tools/list", {}, { "x-api-key": SECRET })).status, 200);
});

test("initialize, then tools/list advertises exactly the five tools", async () => {
  const init = await rpc(`${BASE}/${SECRET}`, "initialize", {
    protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "test", version: "0" },
  });
  assert.equal(init.status, 200);
  assert.equal(init.body.result.serverInfo.name, "come-with-social");
  const r = await rpc(`${BASE}/${SECRET}`, "tools/list");
  const names = r.body.result.tools.map((t: any) => t.name).sort();
  assert.deepEqual(names, ["add_note", "create_post_skeleton", "get_results", "list_posts", "update_post_draft"]);
  assert.ok(!names.some((n: string) => /delete|remove/.test(n)));
  const upd = r.body.result.tools.find((t: any) => t.name === "update_post_draft");
  assert.deepEqual(Object.keys(upd.inputSchema.properties).sort(), ["brief", "claude_caption", "id"]);
  assert.equal(r.body.result.tools.find((t: any) => t.name === "list_posts").annotations.readOnlyHint, true);
});

test("schema refuses junk before any rule runs", async () => {
  const bad = [
    ["list_posts", { from: "Oct 5", to: "2026-10-06" }],
    ["create_post_skeleton", { title: "x", scheduled_at: "tomorrow", account: "come_with", format: "reel", phase: "radio" }],
    ["create_post_skeleton", { title: "x", scheduled_at: "2026-10-06T18:00:00-04:00", account: "instagram", format: "reel", phase: "radio" }],
    ["update_post_draft", { id: "not-a-uuid", claude_caption: "x" }],
    ["update_post_draft", { id: PID, caption: "sneak the final caption in" }],
    ["update_post_draft", { id: PID, stage: "posted" }],
  ] as const;
  for (const [name, args] of bad) {
    const r = await callTool(name, args);
    const failed = r.body.error || r.body.result?.isError;
    assert.ok(failed, `${name} ${JSON.stringify(args)} should fail`);
  }
  assert.equal(s.posts.get(PID)!.caption, "final");
});

test("an approved post is refused through the full stack", async () => {
  const r = await callTool("update_post_draft", { id: PID, claude_caption: "new draft" });
  assert.equal(r.body.result.isError, true);
  assert.match(r.body.result.content[0].text, /only edits posts at idea or drafted/);
  assert.equal(s.posts.get(PID)!.claude_caption, null);
});

test("a delete tool does not exist", async () => {
  const r = await callTool("delete_post", { id: PID });
  assert.ok(r.body.error || r.body.result?.isError);
  assert.ok(s.posts.has(PID));
});

test("a good call returns JSON text content", async () => {
  const r = await callTool("list_posts", { from: "2026-10-01", to: "2026-10-31" });
  assert.equal(r.status, 200);
  assert.ok(!r.body.result.isError);
  const data = JSON.parse(r.body.result.content[0].text);
  assert.equal(data.posts[0].final_caption, "final");
});

test("every call is logged - including ones rejected before a tool runs", async () => {
  LOGS.length = 0;
  await callTool("update_post_draft", { id: PID, caption: "schema says no" });
  await callTool("delete_post", { id: PID });
  await callTool("list_posts", { from: "2026-10-01", to: "2026-10-02" });
  await rpc(`${BASE}/${SECRET}`, "tools/list");
  assert.deepEqual(LOGS.map((l) => [l.tool, l.ok]), [["update_post_draft", false], ["delete_post", false], ["list_posts", true]]);
  assert.equal(LOGS[0].post_id, PID);
  assert.match(LOGS[1].detail, /rejected before the tool ran/);
});
