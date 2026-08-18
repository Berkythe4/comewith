// Tests for the ingest-finance auth gate.
//
//   node --test supabase/functions/ingest-finance/
//
// No network, no database, no real token. The three cases the handoff requires
// (no header → 401, wrong token → 401, correct token → 200) plus the fail-closed
// case, which is the one that actually matters: a server with no PUSH_TOKEN set
// must reject everything rather than wave everything through.
//
// The module calls Deno.serve() at import time, so we stub the Deno global and
// capture the handler before importing it.

import { test } from "node:test";
import assert from "node:assert/strict";

const TOKEN = "push_" + "a".repeat(64);   // shape-accurate, not a real token

// The module reads PUSH_TOKEN per request, not at import, so one import serves
// every case — we just swap what the stubbed env returns.
let serverToken: string | undefined = TOKEN;
let handler: ((req: Request) => Promise<Response>) | null = null;

(globalThis as any).Deno = {
  env: { get: (k: string) => (k === "PUSH_TOKEN" ? serverToken : undefined) },
  serve: (h: (req: Request) => Promise<Response>) => { handler = h; },
};

async function loadHandler(pushToken?: string) {
  serverToken = pushToken;
  if (!handler) await import("./index.ts");
  assert.ok(handler, "handler was not registered");
  return handler!;
}

const post = (headers: Record<string, string> = {}) =>
  new Request("https://example.test/ingest-finance", {
    method: "POST",
    headers: { "Content-Type": "application/json", ...headers },
    body: JSON.stringify({ rows: [{ vendor: "Test Vendor", amount: -12.34 }] }),
  });

test("no Authorization header -> 401", async () => {
  const h = await loadHandler(TOKEN);
  const res = await h(post());
  assert.equal(res.status, 401);
});

test("wrong token -> 401", async () => {
  const h = await loadHandler(TOKEN);
  const res = await h(post({ Authorization: "Bearer push_" + "b".repeat(64) }));
  assert.equal(res.status, 401);
});

test("correct token -> 200 and counts the rows", async () => {
  const h = await loadHandler(TOKEN);
  const res = await h(post({ Authorization: "Bearer " + TOKEN }));
  assert.equal(res.status, 200);
  assert.equal((await res.json()).accepted, 1);
});

test("fails CLOSED when the server has no PUSH_TOKEN set", async () => {
  const h = await loadHandler(undefined);
  const res = await h(post({ Authorization: "Bearer " + TOKEN }));
  assert.equal(res.status, 500, "unset token must reject, never allow");
});

test("a correct token in the query string is still rejected", async () => {
  // Guards the rule that separates this from ingest-email: tokens in URLs end up
  // in access logs, so the query string must never be an accepted channel.
  const h = await loadHandler(TOKEN);
  const res = await h(new Request(
    "https://example.test/ingest-finance?key=" + TOKEN,
    { method: "POST", headers: { "Content-Type": "application/json" }, body: '{"rows":[]}' },
  ));
  assert.equal(res.status, 401);
});

test("auth is checked before the body is parsed", async () => {
  // An unauthenticated caller should not be able to probe body validation.
  const h = await loadHandler(TOKEN);
  const res = await h(new Request("https://example.test/ingest-finance", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: "not json at all",
  }));
  assert.equal(res.status, 401, "malformed body from an anonymous caller must still be 401");
});

test("no auth failure response leaks the expected token or the header", async () => {
  const h = await loadHandler(TOKEN);
  const res = await h(post({ Authorization: "Bearer push_" + "c".repeat(64) }));
  const text = await res.text();
  assert.ok(!text.includes(TOKEN), "response must not contain the expected token");
  assert.ok(!text.includes("ccc"), "response must not echo what was supplied");
});
