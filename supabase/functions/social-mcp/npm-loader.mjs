// Test-only: lets `node --test` load the Deno-style `npm:pkg@ver` imports in
// server.ts by resolving them from ./.test-deps/node_modules. Not deployed
// (deploy_edge_function.py ships *.ts only).
//
//   cd supabase/functions/social-mcp/.test-deps && npm i @modelcontextprotocol/server@2.2.0 zod@4
//   node --import ./supabase/functions/social-mcp/npm-loader.mjs --test supabase/functions/social-mcp/server.test.ts
import { register } from "node:module";

const deps = new URL("./.test-deps/package.json", import.meta.url).href;
const hooks = `
export async function resolve(specifier, context, next) {
  if (specifier.startsWith("npm:")) {
    const bare = specifier.slice(4).replace(/^(@[^/]+\\/[^@/]+|[^@/]+)@[^/]+/, "$1");
    return next(bare, { ...context, parentURL: ${JSON.stringify(deps)} });
  }
  return next(specifier, context);
}`;
register("data:text/javascript," + encodeURIComponent(hooks), import.meta.url);
