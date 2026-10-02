// social-mcp transport: the MCP server (Streamable HTTP, stateless, one fresh
// McpServer per request - Supabase's documented MCP-on-Edge-Functions pattern),
// the zod schema for every tool, and the shared-secret gate.
//
// The secret travels in the URL path - /functions/v1/social-mcp/<secret> -
// because that is what every claude.ai custom connector can send. Request-header
// auth is in beta for a limited set of organisations, so an `x-api-key: <secret>`
// header is accepted as well, for when it reaches this account. Anything without
// the secret gets a bare 404: the endpoint does not admit that it exists.
import { createMcpHandler, McpServer } from "npm:@modelcontextprotocol/server@2.2.0";
import { z } from "npm:zod@4";
import { ACCOUNTS, FORMATS, PHASES, STAGES, runTool, type Store } from "./tools.ts";

const day = z.string().regex(/^\d{4}-\d{2}-\d{2}$/, "YYYY-MM-DD");
const uuid = z.string().uuid();

// .strict() everywhere: an unexpected key (say `caption`, or `stage`) is a
// refusal, never silently dropped.
export const SCHEMAS = {
  list_posts: z.object({
    from: day.describe("First day, inclusive, New York time (YYYY-MM-DD)"),
    to: day.describe("Last day, inclusive, New York time (YYYY-MM-DD)"),
    stage: z.enum(STAGES).optional(),
    account: z.enum(ACCOUNTS).optional(),
  }).strict(),
  create_post_skeleton: z.object({
    title: z.string().min(1).max(200),
    scheduled_at: z.string().datetime({ offset: true }).describe("ISO date-time with offset, e.g. 2026-10-06T18:00:00-04:00"),
    account: z.enum(ACCOUNTS),
    format: z.enum(FORMATS),
    phase: z.enum(PHASES),
    brief: z.string().max(2000).optional(),
  }).strict(),
  update_post_draft: z.object({
    id: uuid,
    brief: z.string().max(2000).optional(),
    claude_caption: z.string().max(5000).optional(),
  }).strict(),
  add_note: z.object({ id: uuid, text: z.string().min(1).max(4000) }).strict(),
  get_results: z.object({ from: day, to: day }).strict(),
};

const DESCRIPTIONS: Record<keyof typeof SCHEMAS, string> = {
  list_posts: "Read-only. Come With social posts scheduled between two dates (New York time), with brief, Claude's caption, the final caption, asset link, results and note count.",
  create_post_skeleton: "Create a planned post. Always lands at stage 'idea', owned by Janelle. Refuses a second post with the same title on the same day.",
  update_post_draft: "Write a brief and/or Claude's caption on a post at stage idea or drafted. Writing a caption moves idea -> drafted. Never touches the final caption; refuses any post at ready or later.",
  add_note: "Add a note to a post's conversation thread, signed 'Claude'.",
  get_results: "Read-only. Posted items between two dates with their results (views, likes, shares, saves) - for the Friday report.",
};

const READ_ONLY = new Set(["list_posts", "get_results"]);

export function buildServer(store: Store): McpServer {
  const server = new McpServer({ name: "come-with-social", version: "1.0.0" });
  for (const name of Object.keys(SCHEMAS) as (keyof typeof SCHEMAS)[]) {
    server.registerTool(name, {
      description: DESCRIPTIONS[name],
      inputSchema: SCHEMAS[name],
      annotations: { readOnlyHint: READ_ONLY.has(name), destructiveHint: false },
    }, async (args: Record<string, unknown>) => {
      const out = await runTool(store, name, args);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(out.ok ? out.result : out, null, 2) }],
        isError: !out.ok,
      };
    });
  }
  return server;
}

// Constant-time compare so the secret cannot be guessed a byte at a time.
export function secretMatches(given: string | null | undefined, secret: string): boolean {
  if (!given || !secret || secret.length < 32) return false;
  const a = new TextEncoder().encode(given), b = new TextEncoder().encode(secret);
  let diff = a.length ^ b.length;
  for (let i = 0; i < b.length; i++) diff |= (a[i % (a.length || 1)] ?? 0) ^ b[i];
  return diff === 0;
}

// The path segment after the function name, e.g.
//   /social-mcp/<secret>        or  /functions/v1/social-mcp/<secret>/mcp
export function pathSecret(url: string): string | null {
  const segs = new URL(url).pathname.split("/").filter(Boolean);
  const i = segs.indexOf("social-mcp");
  return i >= 0 && segs[i + 1] ? decodeURIComponent(segs[i + 1]) : null;
}

const notFound = () => new Response("Not found", { status: 404 });

export function makeFetch(storeFactory: () => Store, secret: string) {
  return async (req: Request): Promise<Response> => {
    const ok = secretMatches(pathSecret(req.url), secret) || secretMatches(req.headers.get("x-api-key"), secret);
    if (!ok) return notFound();
    // One store per request, with its log() watched: a call the SDK rejects
    // before the tool runs (bad schema, an unknown tool such as delete_post)
    // never reaches runTool, and "log every call" includes the refused ones.
    const store = storeFactory();
    let logged = false;
    const watched: Store = { ...store, log: (e) => { logged = true; return store.log(e); } };
    let call: { name: string; id: string | null } | null = null;
    try {
      const msg = await req.clone().json();
      if (msg && msg.method === "tools/call" && msg.params && typeof msg.params.name === "string") {
        const a = msg.params.arguments || {};
        call = { name: String(msg.params.name).slice(0, 80), id: typeof a.id === "string" && /^[0-9a-f-]{36}$/i.test(a.id) ? a.id : null };
      }
    } catch { /* not JSON: the handler answers that itself */ }
    const res = await createMcpHandler(() => buildServer(watched), { responseMode: "json" }).fetch(req);
    if (call && !logged) {
      await store.log({ tool: call.name, post_id: call.id, ok: false, detail: "rejected before the tool ran (input schema or unknown tool)" }).catch(() => {});
    }
    return res;
  };
}
