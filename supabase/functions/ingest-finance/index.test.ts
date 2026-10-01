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
function fakeDb(seed: any[] = [], incomeSeed: any[] = []) {
  const tables: Record<string, any[]> = {
    expenses: [...seed], income: [...incomeSeed], budget_lines: [],
    ingest_queue: [], ingest_runs: [],
  };
  const from = (table: string) => {
    // `single` matters now: the settlement steps need the WHOLE candidate pool
    // to tell "one obvious answer" from "two, so do not choose".
    const st: any = { op: "select", filters: [] as any[], rec: null, patch: null, single: false };
    const match = (rows: any[]) => rows.filter(r =>
      st.filters.every(([c, op, v]: any) => op === "is" ? (r[c] ?? null) === v : r[c] === v));
    const run = async () => {
      const rows = tables[table] ?? (tables[table] = []);
      if (st.op === "insert") { rows.push({ id: "id" + rows.length, ...st.rec }); return { data: null, error: null }; }
      if (st.op === "update") { match(rows).forEach(r => Object.assign(r, st.patch)); return { data: null, error: null }; }
      if (st.op === "delete") { tables[table] = rows.filter(r => !match(rows).includes(r)); return { data: null, error: null }; }
      const m = match(rows);
      return { data: st.single ? (m[0] ?? null) : m, error: null };
    };
    const q: any = {
      select() { st.op = "select"; return q; },
      insert(rec: any) { st.op = "insert"; st.rec = rec; return q; },
      update(patch: any) { st.op = "update"; st.patch = patch; return q; },
      delete() { st.op = "delete"; return q; },
      eq(c: string, v: any) { st.filters.push([c, "eq", v]); return q; },
      is(c: string, v: any) { st.filters.push([c, "is", v]); return q; },
      limit() { return q; },
      maybeSingle: () => { st.single = true; return run(); },
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

async function pushWith(db: any, rows: any[], extra: Record<string, unknown> = {}) {
  await useDb(db);
  const h = await loadHandler(TOKEN);
  return h(new Request("https://example.test/ingest-finance", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: "Bearer " + TOKEN },
    body: JSON.stringify({ rows, ...extra }),
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

test("a row missing its external_ref is QUEUED, not guessed at and not silently dropped", async () => {
  // It used to be counted as `skipped`, which reads as "nothing to do". A row
  // that cannot be deduped on any later run needs a person, so it is queued.
  const db = fakeDb([]);
  const res = await pushWith(db, [{ date: "2026-06-15", amount: 10, kind: "expense" }]);
  const body = await res.json();
  assert.equal(body.queued, 1);
  assert.equal(db.tables.expenses.length, 0, "still must not land on the books");
  assert.equal(db.tables.ingest_queue[0].reason, "missing_ref");
});

test("a gig's revenue and cost budgets coexist in the same month", async () => {
  // Regression: the delete key was (scope, period, category), so the cost row
  // deleted the revenue row it shared a category with. Six budgets vanished.
  const db = fakeDb([]);
  await useDb(db);
  const h = await loadHandler(TOKEN);
  const res = await h(new Request("https://example.test/ingest-finance", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: "Bearer " + TOKEN },
    body: JSON.stringify({
      rows: [],
      budget_lines: [
        { period: "2026-07", category: "DJ Gig #1", planned_amount: 500, direction: "income" },
        { period: "2026-07", category: "DJ Gig #1", planned_amount: 425, direction: "expense" },
      ],
    }),
  }));
  assert.equal((await res.json()).budgets, 2);
  assert.equal(db.tables.budget_lines.length, 2, "both directions must survive");
  const dirs = db.tables.budget_lines.map((b: any) => b.direction).sort();
  assert.deepEqual(dirs, ["expense", "income"]);
});

// ---------------------------------------------------------------------------
// Settle before add.
//
// The cases below are the two that went wrong in production, plus the refusals
// that keep the fix from becoming a new way to be wrong.
// ---------------------------------------------------------------------------

/** A cost incurred at the 2026-08-16 gig, paid through PayPal on 2026-09-08. */
const ACCRUAL = { id: "site-h", date: "2026-08-16", amount: 150, vendor: "Henry",
                  category: "Contractors", status: "paid", event_id: "ev-816",
                  external_ref: null, deleted_at: null, ledger: "come_with" };
const PAYMENT = { external_ref: "pp-150", date: "2026-09-08", amount: 150, kind: "expense",
                  category: "Operations", vendor: "Henry Zaradich", funded_by: "business",
                  cash_source: "paypal", ledger: "come_with" };

test("S3: a payment SETTLES the cost it belongs to instead of inserting beside it", async () => {
  // The $250 bug: same obligation, different dates, counted twice in two months.
  const db = fakeDb([{ ...ACCRUAL }]);
  const body = await (await pushWith(db, [PAYMENT])).json();
  assert.equal(body.settled, 1, "should settle the August row");
  assert.equal(body.inserted, 0, "must NOT insert a September duplicate");
  assert.equal(db.tables.expenses.length, 1);
});

test("S3: settling keeps the cost in the month it was INCURRED", async () => {
  const db = fakeDb([{ ...ACCRUAL }]);
  await pushWith(db, [PAYMENT]);
  const row = db.tables.expenses[0];
  assert.equal(row.date, "2026-08-16", "the P&L date must not move to the payment date");
  assert.equal(row.settled_at, "2026-09-08", "the cash date lands in settled_at");
  assert.equal(row.event_id, "ev-816", "event attribution survives");
  assert.equal(row.cash_source, "paypal", "but where the money moved comes from Jennifer");
  assert.equal(row.external_ref, "pp-150", "and the row is now claimed");
});

test("S1: an open payable is settled and marked paid", async () => {
  const db = fakeDb([{ id: "site-a", date: "2026-08-20", amount: 900, vendor: "Berky",
                       status: "accrued", external_ref: null, deleted_at: null,
                       ledger: "come_with" }]);
  const body = await (await pushWith(db, [{ external_ref: "pp-900", date: "2026-09-10",
      amount: 900, kind: "expense", vendor: "Berky", cash_source: "paypal" }])).json();
  assert.equal(body.settled, 1);
  assert.equal(db.tables.expenses[0].status, "paid", "accrued -> paid on settlement");
  assert.equal(db.tables.expenses[0].date, "2026-08-20", "still incurred in August");
});

test("two rows could be the payment -> QUEUED, and nothing is chosen", async () => {
  // Real case: two identical $100 Keith Berkman rows on the same day.
  const db = fakeDb([
    { id: "a", date: "2026-08-16", amount: 100, vendor: "Berky", status: "paid",
      external_ref: null, deleted_at: null, ledger: "come_with" },
    { id: "b", date: "2026-08-17", amount: 100, vendor: "Berky", status: "paid",
      external_ref: null, deleted_at: null, ledger: "come_with" },
  ]);
  const body = await (await pushWith(db, [{ external_ref: "pp-100", date: "2026-09-08",
      amount: 100, kind: "expense", vendor: "Berky", cash_source: "paypal" }])).json();
  assert.equal(body.queued, 1);
  assert.equal(body.settled, 0);
  assert.equal(body.inserted, 0, "ambiguity must not fall through to insert");
  assert.equal(db.tables.ingest_queue[0].reason, "ambiguous_settlement");
  assert.equal(db.tables.ingest_queue[0].candidate_ids.length, 2);
});

test("the amount and timing fit but the payee does not -> QUEUED", async () => {
  const db = fakeDb([{ id: "site-x", date: "2026-08-16", amount: 100, vendor: "Berky",
                       status: "paid", external_ref: null, deleted_at: null,
                       ledger: "come_with" }]);
  const body = await (await pushWith(db, [{ external_ref: "pp-x", date: "2026-09-08",
      amount: 100, kind: "expense", vendor: "Keith Berkman", cash_source: "paypal" }])).json();
  assert.equal(body.queued, 1);
  assert.equal(db.tables.ingest_queue[0].reason, "vendor_mismatch");
  assert.equal(db.tables.expenses[0].external_ref, null, "the site row is untouched");
});

test("a payment older than the window does not reach back and claim something", async () => {
  const db = fakeDb([{ id: "old", date: "2026-01-01", amount: 150, vendor: "Henry",
                       status: "paid", external_ref: null, deleted_at: null,
                       ledger: "come_with" }]);
  const body = await (await pushWith(db, [PAYMENT])).json();
  assert.equal(body.settled, 0, "8 months is beyond the 90-day window");
  assert.equal(body.inserted, 1, "so it is genuinely new");
});

test("a payment never settles a row on the other ledger", async () => {
  const db = fakeDb([{ id: "di", date: "2026-08-16", amount: 150, vendor: "Henry",
                       status: "paid", external_ref: null, deleted_at: null,
                       ledger: "dance_infusion" }]);
  const body = await (await pushWith(db, [PAYMENT])).json();
  assert.equal(body.settled, 0, "Dance Infusion must never absorb a Come With payment");
  assert.equal(body.inserted, 1);
});

test("S0: an update never drags an event-linked cost to the payment date", async () => {
  const db = fakeDb([{ ...ACCRUAL }]);
  await pushWith(db, [PAYMENT]);            // settles it, date 2026-08-16
  await pushWith(db, [PAYMENT]);            // now it is ours -> S0 update
  assert.equal(db.tables.expenses[0].date, "2026-08-16", "the accrual date must survive re-pushes");
  assert.equal(db.tables.expenses.length, 1);
});

test("report_only computes the same answer and writes NOTHING", async () => {
  const db = fakeDb([{ ...ACCRUAL }]);
  const body = await (await pushWith(db, [PAYMENT], { report_only: true })).json();
  assert.equal(body.report_only, true);
  assert.equal(body.settled, 1, "it still says what it WOULD do");
  assert.equal(db.tables.expenses[0].external_ref, null, "but claims nothing");
  assert.equal(db.tables.expenses[0].settled_at, undefined, "and settles nothing");
  assert.equal(db.tables.ingest_runs.length, 1, "the report itself is recorded");
  assert.equal(db.tables.ingest_runs[0].report_only, true);
});

test("every run is recorded so the site can answer 'is this current?'", async () => {
  const db = fakeDb([]);
  await pushWith(db, [ROW], { source: "jennifer" });
  const run = db.tables.ingest_runs[0];
  assert.equal(run.source, "jennifer");
  assert.equal(run.inserted, 1);
  assert.equal(run.report_only, false);
});
