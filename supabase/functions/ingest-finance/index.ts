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
// DEPLOY: this is NOT live until explicitly deployed —
//   python scripts/deploy_edge_function.py ingest-finance
// It must be deployed with JWT verification OFF (--no-verify-jwt): Jennifer
// authenticates with the push token, not a Supabase user JWT.

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

  const rows = (body as { rows?: unknown[] })?.rows;
  if (!Array.isArray(rows)) return err(400, "expected { rows: [...] }");
  if (rows.length > 5000) return err(413, "too many rows in one push");

  // Persistence is deliberately not wired yet: the payload contract between
  // Jennifer and this repo is still being agreed, and landing a table here would
  // mean a migration that has not been introspected or dry-run against prod
  // (CLAUDE.md). Auth is the piece the handoff asked for, and it is complete —
  // this accepts and counts, so the token path can be verified end to end before
  // any schema exists.
  return ok({ accepted: rows.length, stored: 0, note: "auth verified; storage pending payload contract" });
});
