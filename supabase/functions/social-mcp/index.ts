// social-mcp  (remote MCP server for claude.ai - SECRET-gated, no login)
//
// Gives Claude a narrow, safe handle on the social calendar so it can plan and
// draft from claude.ai and from scheduled tasks. Five tools, no deletes, never
// the final caption, never a post at `approved` or later. See tools.ts for the
// rules and server.ts for the transport + secret.
//
// Deployed with verify_jwt = false (claude.ai sends no Supabase JWT); the
// SOCIAL_MCP_SECRET path segment is the credential. Rotate it by setting a new
// secret and re-adding the connector in claude.ai with the new URL.
//
// Env: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY (platform-provided),
//      SOCIAL_MCP_SECRET (>= 32 chars; unset = every request 404s).
import { makeFetch } from "./server.ts";
import { supabaseStore } from "./store.ts";

const url = Deno.env.get("SUPABASE_URL") ?? "";
const key = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const secret = Deno.env.get("SOCIAL_MCP_SECRET") ?? "";

const handle = makeFetch(() => supabaseStore(url, key), secret);
Deno.serve((req) => handle(req));
