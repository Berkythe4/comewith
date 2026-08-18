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
  await useDb(fakeDb([]));
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

// ---------------------------------------------------------------------------
// Storage: insert vs ADOPT vs update.
//
// Adopt is the behaviour that makes the true-up safe. Measured against prod:
// Jennifer holds 180 Come With rows, this database holds 133 expenses, and 66
// are the SAME charge in both. Without adopt, the first push creates 66
// duplicates and every P&L number after that is wrong.
// ---------------------------------------------------------------------------

/** Minimal stand-in for supabase-js: chainable AND awaitable, like the real one. */
function fakeDb(seed: any[] = []) {
  const tables: Record<string, any[]> = { expenses: [...seed], income: [], budget_lines: [] };
  const from = (table: string) => {
    const st: any = { op: "select", filters: [] as any[], rec: null, patch: null };
    const match = (rows: any[]) => rows.filter(r =>
      st.filters.every(([c, op, v]: any) => op === "is" ? (r[c] ?? null) === v : r[c] === v));
    const run = async () => {
      const rows = tables[table];
      if (st.op === "insert") { rows.push({ id: "id" + rows.length, ...st.rec }); return { data: null, error: null }; }
      if (st.op === "update") { match(rows).forEach(r => Object.assign(r, st.patch)); return { data: null, error: null }; }
      if (st.op === "delete") { const gone = match(rows); tables[table] = rows.filter(r => !gone.includes(r)); return { data: null, error: null }; }
      const m = match(rows); return { data: m[0] ?? null, error: null };
    };
    const q: any = {
      select() { st.op = "select"; return q; },
      insert(rec: any) { st.op = "insert"; st.rec = rec; return q; },
      update(patch: any) { st.op = "update"; st.patch = patch; return q; },
      delete() { st.op = "delete"; return q; },
      eq(c: string, v: any) { st.filters.push([c, "eq", v]); return q; },
      is(c: string, v: any) { st.filters.push([c, "is", v]); return q; },
      limit() { return q; },
      maybeSingle: () => run(),
      then: (res: any, rej: any) => run().then(res, rej),   // awaitable
    };
    return q;
  };
  return { client: { from }, tables };
}

/** Point the module at a fake before any test runs, so nothing reaches npm:. */
async function useDb(db: any) {
  await loadHandler(TOKEN);                       // ensures the module is imported
  const mod: any = await import("./index.ts");
  mod.__setDbFactory(async () => db.client);
}

async function pushWith(db: any, rows: any[]) {
  await useDb(db);
  const h = await loadHandler(TOKEN);
  return h(new Request("https://example.test/ingest-finance", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: "Bearer " + TOKEN },
    body: JSON.stringify({ rows }),
  }));
}

const ROW = {
  external_ref: "hash-abc", date: "2026-06-15", amount: 200, kind: "expense",
  category: "Software", vendor: "Splice", funded_by: "owner",
};

test("a charge the site has never seen is INSERTED", async () => {
  const db = fakeDb([]);
  const res = await pushWith(db, [ROW]);
  const body = await res.json();
  assert.equal(res.status, 200);
  assert.equal(body.inserted, 1);
  assert.equal(body.adopted, 0);
  assert.equal(db.tables.expenses.length, 1);
});

test("a hand-entered row with the same date+amount is ADOPTED, not duplicated", async () => {
  // This is the 66-row case. The site's own row has no external_ref.
  const db = fakeDb([{ id: "site-1", date: "2026-06-15", amount: 200,
                       category: "Operations", vendor: "Typed by hand",
                       external_ref: null, deleted_at: null }]);
  const res = await pushWith(db, [ROW]);
  const body = await res.json();
  assert.equal(body.adopted, 1, "should adopt the existing row");
  assert.equal(body.inserted, 0, "must NOT insert a duplicate");
  assert.equal(db.tables.expenses.length, 1, "still exactly one row for this charge");
  assert.equal(db.tables.expenses[0].external_ref, "hash-abc", "row is now claimed");
});

test("adopting preserves the site's own curation", async () => {
  const db = fakeDb([{ id: "site-1", date: "2026-06-15", amount: 200,
                       category: "Operations", vendor: "Typed by hand",
                       event_id: "ev-9", external_ref: null, deleted_at: null }]);
  await pushWith(db, [ROW]);
  const row = db.tables.expenses[0];
  assert.equal(row.category, "Operations", "hand-set category must survive");
  assert.equal(row.vendor, "Typed by hand", "hand-set vendor must survive");
  assert.equal(row.event_id, "ev-9", "event link must survive");
  assert.equal(row.funded_by, "owner", "but funding source is taken from Jennifer");
});

test("re-sending the same file changes nothing (idempotent)", async () => {
  const db = fakeDb([]);
  await pushWith(db, [ROW]);
  const res = await pushWith(db, [ROW]);
  const body = await res.json();
  assert.equal(body.updated, 1, "second send updates in place");
  assert.equal(body.inserted, 0);
  assert.equal(db.tables.expenses.length, 1, "still one row after two pushes");
});

test("a row missing its external_ref is skipped, not guessed at", async () => {
  const db = fakeDb([]);
  const res = await pushWith(db, [{ date: "2026-06-15", amount: 10, kind: "expense" }]);
  const body = await res.json();
  assert.equal(body.skipped, 1);
  assert.equal(db.tables.expenses.length, 0);
});
