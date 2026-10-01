// ingest-finance
//
// Receives fee / vendor rows pushed from Jennifer (Keith's local planner) so the
// Come With P&L can live here rather than on his machine. One client, one
// endpoint, write-only.
//
// SETTLE BEFORE ADD. The function used to ask one question of an incoming
// payment — "is there a row with this exact date and amount?" — and insert if
// the answer was no. That is blind to the thing Come With actually does: a cost
// is incurred at the gig and paid weeks later, on a different date. Two real
// failures from 2026:
//
//   Henry $150 and Berky $100 were recorded against the 2026-08-16 event and
//   paid through PayPal on 2026-09-08. Same-date adopt could not see the August
//   rows, so it inserted its own. $250 of contractor cost counted twice, in two
//   different months.
//
//   A $357.39 Stripe payout was the NET of a $361.00 production fee already on
//   these books (less a $3.61 platform fee). Treated as new money it
//   double-counted revenue while the cash figure still looked right.
//
// So each row now walks an ordered hierarchy and takes the FIRST step that
// fits. See STEPS below. Where more than one row could be the answer it refuses
// to choose and writes to `ingest_queue` instead — a guess here is silent and
// lands in the P&L, which is exactly the class of error that took a month to
// notice.
//
// WHAT SETTLEMENT MUST NOT DO: overwrite the site's own `date`. Jennifer sends
// the date the cash moved; this database wants the date the cost was incurred,
// with `settled_at` carrying the cash date (177). Clobbering `date` would drag
// every settled cost into the month it was paid and undo the accrual.
//
// SECURITY (see HANDOFF-push-token-security.md — implemented as written there):
//   - Static bearer token in the Authorization header. NEVER a ?key= query
//     param: query strings land in access logs and browser history. Note this
//     deliberately differs from `ingest-email`, which does use ?key= — that one
//     predates this rule and is constrained by what inbound mail providers can
//     send. Do not copy that pattern into new functions.
//   - Constant-time comparison, on SHA-256 digests so the two sides are always
//     the same length and no timing signal leaks from an early mismatch.
//   - FAIL CLOSED: if the server's own PUSH_TOKEN is unset we return 500 and
//     accept nothing. "No token configured" must never mean "auth disabled".
//   - Bare 401 for every auth failure. We never distinguish missing-header from
//     wrong-token; that difference is free reconnaissance.
//   - The Authorization header is never logged, echoed, or included in an error
//     body. Debug from the status code.
//
// Secret: PUSH_TOKEN (project secret — `supabase secrets set PUSH_TOKEN=...`).
// Rotation: docs/ROTATE_PUSH_TOKEN.md
//
// STORAGE: writes to `expenses` / `income` keyed on external_ref, replaces
// period budget_lines, and records every run in `ingest_runs` (migration 207).
// Uses the service role, so it enforces its own auth — the bearer check above is
// the only gate.
//
// DEPLOY: this is NOT live until explicitly deployed —
//   python scripts/deploy_edge_function.py ingest-finance
// It must be deployed with JWT verification OFF (--no-verify-jwt): Jennifer
// authenticates with the push token, not a Supabase user JWT.

type Row = {
  external_ref?: string; date?: string; amount?: number; kind?: string;
  category?: string; vendor?: string; description?: string;
  funded_by?: string; event_na?: boolean;
  cash_source?: string; ledger?: string; settled_at?: string;
};
type BudgetLine = {
  period?: string; category?: string; planned_amount?: number;
  direction?: string; notes?: string;
};
type Payload = {
  rows?: Row[]; budget_lines?: BudgetLine[];
  // A dry run: compute every decision, write nothing but the run record. The
  // settlement rules reach backwards over months of history, so they get looked
  // at before they get applied.
  report_only?: boolean;
  source?: string;
};

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const ok = (o: unknown) => new Response(JSON.stringify(o), { headers: JH });
const err = (s: number, m: string) =>
  new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const sha256 = async (s: string): Promise<Uint8Array> =>
  new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)));

// Constant-time byte compare. Both inputs are SHA-256 digests, so lengths always
// match and the loop always runs to completion regardless of where they differ.
function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

/** True only for a well-formed header carrying the right token. */
async function authorized(req: Request): Promise<boolean> {
  const expected = Deno.env.get("PUSH_TOKEN");
  if (!expected) throw new Error("unconfigured");   // -> 500, fail closed
  const provided = req.headers.get("Authorization") ?? "";
  if (!provided.startsWith("Bearer ")) return false;
  const token = provided.slice(7);
  if (!token) return false;
  return timingSafeEqual(await sha256(token), await sha256(expected));
}

// The Supabase client is reached through this indirection so the settlement
// logic below can be tested against a fake. Production path is unchanged: a lazy
// import AFTER auth, so an unauthenticated caller never loads the client.
type DbFactory = () => Promise<any>;
let dbFactory: DbFactory = async () => {
  const { createClient } = await import("npm:@supabase/supabase-js@2");
  return createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    { auth: { persistSession: false } },
  );
};
const getDb = () => dbFactory();
/** Test seam. Not used in production. */
export function __setDbFactory(f: DbFactory) { dbFactory = f; }

// ---------------------------------------------------------------------------
// Settlement
// ---------------------------------------------------------------------------

// How far back a payment may reach to settle an obligation. Come With books a
// gig, plays it, and gets paid on the client's terms; 90 days covers that
// without letting a payment claim something from last season.
const SETTLE_WINDOW_DAYS = 90;

const CASH = ["paypal", "bank", "personal", "other"];

function days(a?: string, b?: string): number {
  if (!a || !b) return Number.POSITIVE_INFINITY;
  const ms = Date.parse(b) - Date.parse(a);
  return Number.isNaN(ms) ? Number.POSITIVE_INFINITY : ms / 86400000;
}

/** Loose vendor comparison. Used only to CONFIRM a match, never to make one. */
function vendorAgrees(candidate: any, r: Row): boolean {
  const a = (candidate?.vendor ?? "").trim().toLowerCase();
  const b = (r.vendor ?? "").trim().toLowerCase();
  if (!a || !b) return false;
  if (a === b) return true;
  // "Henry" vs "Henry Zaradich" — a shared word of real length is enough. Two
  // payees who share only "the" or "llc" are not the same person, so one-and
  // two-letter tokens do not count.
  const tok = (s: string) => s.split(/[^a-z0-9]+/).filter(w => w.length > 2);
  const ta = tok(a), tb = tok(b);
  return ta.some(w => tb.includes(w));
}

/**
 * Place ONE incoming row. Returns what happened and why; performs no writes
 * when `dry` is set.
 *
 * THE STEPS, in order — the first that fits wins:
 *   S0 identity    external_ref is already ours          -> update in place
 *   S1 payable     an open accrued/invoiced obligation   -> settle it
 *   S2 same-day    an unreffed row, same date + amount   -> adopt (the old rule)
 *   S3 near-date   an unreffed row, same amount, in window -> settle it
 *   S4 new         nothing matched                       -> insert
 * Ambiguity at S1/S3 never falls through to S4. It queues.
 */
async function placeRow(db: any, r: Row, dry: boolean) {
  const table = r.kind === "income" ? "income" : "expenses";
  const ledger = r.ledger === "dance_infusion" ? "dance_infusion" : "come_with";

  const rec: Record<string, unknown> = {
    external_ref: r.external_ref,
    date: r.date,
    amount: r.amount,
    category: r.category ?? null,
    description: r.description ?? null,
  };
  if (table === "expenses") {
    rec.vendor = r.vendor ?? null;
    rec.funded_by = r.funded_by === "owner" ? "owner" : "business";
    rec.event_na = r.event_na !== false;
    // Only the sender knows where the money physically moved; an unrecognised
    // value becomes null rather than a guess, because this drives the cash float.
    if (CASH.indexOf(r.cash_source ?? "") >= 0) rec.cash_source = r.cash_source;
    if (r.ledger === "dance_infusion") rec.ledger = "dance_infusion";
  } else if (CASH.indexOf(r.cash_source ?? "") >= 0) {
    rec.cash_source = r.cash_source;
  }

  // ---- S0. Already ours -> update in place. -------------------------------
  const { data: mine } = await db.from(table).select("id,date,event_id,settled_at")
    .eq("external_ref", r.external_ref).maybeSingle();
  if (mine) {
    const patch = { ...rec };
    // DO NOT drag a settled or event-linked cost into the month it was paid.
    // Jennifer sends the cash date; this row's `date` is when it was incurred
    // and belongs to the site. See the header.
    if (mine.event_id || mine.settled_at) delete patch.date;
    if (!dry) {
      const { error } = await db.from(table).update(patch).eq("id", mine.id);
      if (error) return { action: "problem", detail: error.message };
    }
    return { action: "updated", id: mine.id, kept_date: !!(mine.event_id || mine.settled_at) };
  }

  // Candidate pool: same amount, never claimed, same ledger, not deleted. Date
  // and status are filtered here rather than in the query so the shape stays
  // simple enough to be faked in tests and cheap on PostgREST.
  const { data: poolRaw } = await db.from(table)
    .select("id,date,amount,vendor,status,event_id,settled_at,ledger,expected_amount")
    .eq("amount", r.amount).is("external_ref", null).is("deleted_at", null);
  const pool: any[] = (Array.isArray(poolRaw) ? poolRaw : poolRaw ? [poolRaw] : [])
    .filter(c => (c.ledger ?? "come_with") === ledger);

  const inWindow = (c: any) => {
    const d = days(c.date, r.date);          // obligation first, payment after
    return d >= 0 && d <= SETTLE_WINDOW_DAYS;
  };
  const settledPatch = (c: any) => {
    const p: Record<string, unknown> = {
      external_ref: r.external_ref,
      status: table === "expenses" ? "paid" : "received",
      settled_at: r.settled_at ?? r.date,
    };
    // Where the money moved is exactly what this database could not know on its
    // own. Category, vendor, date and event link are its own curation — kept.
    if (rec.cash_source) p.cash_source = rec.cash_source;
    if (table === "expenses") p.funded_by = rec.funded_by;
    return p;
  };
  const queue = (reason: string, cands: any[], detail?: unknown) =>
    ({ action: "queued", reason, candidate_ids: cands.map(c => c.id), detail });

  // ---- S1. An open obligation this payment discharges. --------------------
  const open = pool.filter(c => (c.status === "accrued" || c.status === "invoiced") && inWindow(c));
  if (open.length === 1 && vendorAgrees(open[0], r)) {
    if (!dry) {
      const { error } = await db.from(table).update(settledPatch(open[0])).eq("id", open[0].id);
      if (error) return { action: "problem", detail: error.message };
    }
    return { action: "settled", step: "S1", id: open[0].id };
  }
  if (open.length > 1) return queue("ambiguous_settlement", open);
  if (open.length === 1) return queue("vendor_mismatch", open,
    { site_vendor: open[0].vendor ?? null, incoming_vendor: r.vendor ?? null });

  // ---- S2. The old rule: an unreffed row on the very same day. ------------
  const sameDay = pool.filter(c => c.date === r.date);
  if (sameDay.length === 1) {
    const claim: Record<string, unknown> = { external_ref: r.external_ref };
    if (table === "expenses") claim.funded_by = rec.funded_by;
    if (rec.cash_source) claim.cash_source = rec.cash_source;
    if (!dry) {
      const { error } = await db.from(table).update(claim).eq("id", sameDay[0].id);
      if (error) return { action: "problem", detail: error.message };
    }
    return { action: "adopted", step: "S2", id: sameDay[0].id };
  }
  if (sameDay.length > 1) return queue("ambiguous_settlement", sameDay);

  // ---- S3. Same amount, earlier date, inside the window. ------------------
  // This is the accrual case: incurred in August, paid in September.
  const near = pool.filter(inWindow);
  if (near.length === 1 && vendorAgrees(near[0], r)) {
    if (!dry) {
      const { error } = await db.from(table).update(settledPatch(near[0])).eq("id", near[0].id);
      if (error) return { action: "problem", detail: error.message };
    }
    return { action: "settled", step: "S3", id: near[0].id };
  }
  if (near.length > 1) return queue("ambiguous_settlement", near);
  if (near.length === 1) return queue("vendor_mismatch", near,
    { site_vendor: near[0].vendor ?? null, incoming_vendor: r.vendor ?? null });

  // ---- S4. Genuinely new. -------------------------------------------------
  if (!dry) {
    const { error } = await db.from(table).insert(rec);
    if (error) return { action: "problem", detail: error.message };
  }
  return { action: "inserted", step: "S4" };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "method not allowed");

  try {
    if (!(await authorized(req))) return err(401, "unauthorized");
  } catch (e) {
    // Only reachable when PUSH_TOKEN is missing from the environment. Say that
    // the server is misconfigured; never hint at what was supplied.
    console.error("ingest-finance: PUSH_TOKEN is not set — refusing all requests");
    return err(500, "server not configured");
  }

  let body: unknown;
  try {
    body = await req.json();
  } catch {
    return err(400, "body must be JSON");
  }

  const rows = (body as Payload)?.rows;
  const budgetLines = (body as Payload)?.budget_lines ?? [];
  const dry = (body as Payload)?.report_only === true;
  const source = (body as Payload)?.source ?? null;
  if (!Array.isArray(rows)) return err(400, "expected { rows: [...] }");
  if (rows.length > 5000) return err(413, "too many rows in one push");

  const db = await getDb();

  const tally = { settled: 0, inserted: 0, updated: 0, adopted: 0, queued: 0, skipped: 0 };
  const problems: string[] = [];
  const detail: any[] = [];

  for (const r of rows) {
    if (!r?.date || typeof r.amount !== "number") { tally.skipped++; continue; }
    // A row with no external_ref cannot be deduped on any later run, so it is
    // queued rather than skipped silently — an invisible skip is how rows go
    // missing without anyone being told.
    if (!r.external_ref) {
      tally.queued++;
      const q = {
        kind: r.kind === "income" ? "income" : "expense", date: r.date, amount: r.amount,
        vendor: r.vendor ?? null, category: r.category ?? null, description: r.description ?? null,
        cash_source: r.cash_source ?? null, funded_by: r.funded_by ?? null,
        ledger: r.ledger ?? "come_with", reason: "missing_ref", candidate_ids: [],
      };
      detail.push({ ...q, action: "queued" });
      if (!dry) { const { error } = await db.from("ingest_queue").insert(q); if (error) problems.push(error.message); }
      continue;
    }

    let out: any;
    try {
      out = await placeRow(db, r, dry);
    } catch (e) {
      problems.push(`${r.external_ref}: ${(e as Error).message}`);
      continue;
    }

    if (out.action === "problem") { problems.push(out.detail); continue; }
    if (out.action === "queued") {
      tally.queued++;
      const q = {
        external_ref: r.external_ref, kind: r.kind === "income" ? "income" : "expense",
        date: r.date, amount: r.amount, vendor: r.vendor ?? null, category: r.category ?? null,
        description: r.description ?? null, cash_source: r.cash_source ?? null,
        funded_by: r.funded_by ?? null, ledger: r.ledger ?? "come_with",
        reason: out.reason, candidate_ids: out.candidate_ids ?? [],
        detail: out.detail ?? null,
      };
      detail.push({ ...q, action: "queued" });
      if (!dry) { const { error } = await db.from("ingest_queue").insert(q); if (error) problems.push(error.message); }
      continue;
    }

    (tally as any)[out.action]++;
    detail.push({
      external_ref: r.external_ref, date: r.date, amount: r.amount, vendor: r.vendor ?? null,
      action: out.action, step: out.step ?? "S0", id: out.id ?? null,
      kept_date: out.kept_date ?? undefined,
    });
  }

  // Period budgets are replaced wholesale per (period, category): Jennifer owns
  // the Come With budget until the site grows its own editor for it.
  let budgets = 0;
  if (!dry) {
    for (const b of budgetLines) {
      if (!b?.period || !b?.category) continue;
      // Key MUST include direction: a gig has a revenue budget AND a cost budget in
      // the same month under the same category, so deleting on (period, category)
      // alone made the cost row wipe the revenue row. Six gig revenue budgets went
      // missing exactly this way on the first push.
      await db.from("budget_lines").delete()
        .eq("scope", "period").eq("period", b.period).eq("category", b.category)
        .eq("direction", b.direction === "income" ? "income" : "expense");
      const { error } = await db.from("budget_lines").insert({
        scope: "period", period: b.period, category: b.category,
        planned_amount: b.planned_amount ?? 0,
        direction: b.direction === "income" ? "income" : "expense",
        notes: b.notes ?? null,
      });
      if (!error) budgets++;
    }
  } else {
    budgets = budgetLines.filter(b => b?.period && b?.category).length;
  }

  const summary = {
    accepted: rows.length, ...tally, budgets,
    report_only: dry,
    problems: problems.slice(0, 10),
  };

  // The run record is written even for a dry run — knowing a report was taken,
  // and what it said, is the point of taking it. It is also how the dashboard
  // answers "is this current?" for a job that runs on a laptop it cannot see.
  const { error: runErr } = await db.from("ingest_runs").insert({
    source, accepted: rows.length, settled: tally.settled, inserted: tally.inserted,
    updated: tally.updated, adopted: tally.adopted, queued: tally.queued,
    skipped: tally.skipped, budgets, report_only: dry,
    detail: detail.slice(0, 500), problems: problems.slice(0, 10),
  });
  if (runErr) summary.problems = [...summary.problems, runErr.message].slice(0, 10);

  return ok({ ...summary, detail: detail.slice(0, 500) });
});
