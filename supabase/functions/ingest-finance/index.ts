// ingest-finance
//
// Receives fee / vendor rows pushed from Jennifer (Keith's local planner) so the
// Come With P&L can live here rather than on his machine. One client, one
// endpoint, write-only.
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
// STORAGE: writes to `expenses` / `income` keyed on external_ref, and replaces
// period budget_lines. Requires migration 147 (external_ref, funded_by,
// budget_lines.period). Uses the service role, so it enforces its own auth — the
// bearer check above is the only gate.
//
// DEPLOY: this is NOT live until explicitly deployed —
//   python scripts/deploy_edge_function.py ingest-finance
// It must be deployed with JWT verification OFF (--no-verify-jwt): Jennifer
// authenticates with the push token, not a Supabase user JWT.

type Row = {
  external_ref?: string; date?: string; amount?: number; kind?: string;
  category?: string; vendor?: string; description?: string;
  funded_by?: string; event_na?: boolean;
};
type BudgetLine = {
  period?: string; category?: string; planned_amount?: number;
  direction?: string; notes?: string;
};
type Payload = { rows?: Row[]; budget_lines?: BudgetLine[] };

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

// The Supabase client is reached through this indirection so the adopt/insert
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
  if (!Array.isArray(rows)) return err(400, "expected { rows: [...] }");
  if (rows.length > 5000) return err(413, "too many rows in one push");

  const db = await getDb();

  let inserted = 0, updated = 0, adopted = 0, skipped = 0;
  const problems: string[] = [];

  for (const r of rows) {
    if (!r?.external_ref || !r?.date || typeof r.amount !== "number") { skipped++; continue; }
    const table = r.kind === "income" ? "income" : "expenses";
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
    }

    // 1. Already ours? Update in place.
    const { data: mine } = await db.from(table).select("id")
      .eq("external_ref", r.external_ref).maybeSingle();
    if (mine) {
      const { error } = await db.from(table).update(rec).eq("id", (mine as any).id);
      error ? problems.push(error.message) : updated++;
      continue;
    }

    // 2. ADOPT. The site kept its own record of this spending by hand long before
    //    Jennifer pushed anything — 66 of these charges exist in both places. A
    //    same-date, same-amount row with no external_ref is that same charge, so
    //    we claim it rather than inserting a duplicate beside it. This is the one
    //    behaviour that makes the true-up safe to run.
    const { data: orphan } = await db.from(table).select("id")
      .eq("date", r.date).eq("amount", r.amount).is("external_ref", null)
      .is("deleted_at", null).limit(1).maybeSingle();
    if (orphan) {
      // Only stamp identity + funding. The hand-entered category, vendor and
      // event link are the site's own curation and are deliberately preserved.
      const claim: Record<string, unknown> = { external_ref: r.external_ref };
      if (table === "expenses") claim.funded_by = rec.funded_by;
      const { error } = await db.from(table).update(claim).eq("id", (orphan as any).id);
      error ? problems.push(error.message) : adopted++;
      continue;
    }

    // 3. Genuinely new.
    const { error } = await db.from(table).insert(rec);
    error ? problems.push(error.message) : inserted++;
  }

  // Period budgets are replaced wholesale per (period, category): Jennifer owns
  // the Come With budget until the site grows its own editor for it.
  let budgets = 0;
  for (const b of budgetLines) {
    if (!b?.period || !b?.category) continue;
    await db.from("budget_lines").delete()
      .eq("scope", "period").eq("period", b.period).eq("category", b.category);
    const { error } = await db.from("budget_lines").insert({
      scope: "period", period: b.period, category: b.category,
      planned_amount: b.planned_amount ?? 0,
      direction: b.direction === "income" ? "income" : "expense",
      notes: b.notes ?? null,
    });
    if (!error) budgets++;
  }

  return ok({
    accepted: rows.length,
    inserted, updated, adopted, skipped, budgets,
    problems: problems.slice(0, 10),
  });
});
