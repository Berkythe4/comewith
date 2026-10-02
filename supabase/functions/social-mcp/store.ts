// social-mcp data layer. Service role, server-side only - the key never leaves
// the function. It is scoped by CODE to social_posts, social_post_notes, the
// one profiles lookup that finds Janelle, and connector_log. There is no
// delete anywhere in this file, and the only update is the conditional one.
import { createClient } from "npm:@supabase/supabase-js@2";
import type { Post, Store } from "./tools.ts";

const COLS = "id,title,scheduled_for,posted_at,account,format,phase,stage,brief,claude_caption," +
  "claude_drafted_at,caption,asset_url,results,deleted_at,notes:social_post_notes(count)";

const flat = (r: any): Post => {
  const { notes, ...p } = r;
  return { ...p, note_count: (notes && notes[0] && notes[0].count) || 0 };
};
const must = <T>(res: { data: T; error: any }): T => {
  if (res.error) throw new Error(res.error.message);
  return res.data;
};
// A hard ceiling with a stated reason, not a silent one: 120 days of posts is a
// few hundred rows at most, and PostgREST truncates at 1000 without saying so.
const LIMIT = 900;

export function supabaseStore(url: string, serviceKey: string): Store {
  const sb = createClient(url, serviceKey, { auth: { persistSession: false } });
  return {
    async listPosts({ fromIso, toIso, stage, account }) {
      let q = sb.from("social_posts").select(COLS).is("deleted_at", null)
        .gte("scheduled_for", fromIso).lt("scheduled_for", toIso);
      if (stage) q = q.eq("stage", stage);
      if (account) q = q.eq("account", account);
      const rows = must(await q.order("scheduled_for").order("id").limit(LIMIT)) as any[];
      if (rows.length >= LIMIT) throw new Error(`more than ${LIMIT} posts in range - ask for a shorter range`);
      return rows.map(flat);
    },
    async listPosted({ fromIso, toIso }) {
      const rows = must(await sb.from("social_posts").select(COLS).is("deleted_at", null).eq("stage", "posted")
        .or(`and(posted_at.gte.${fromIso},posted_at.lt.${toIso}),` +
            `and(posted_at.is.null,scheduled_for.gte.${fromIso},scheduled_for.lt.${toIso})`)
        .order("id").limit(LIMIT)) as any[];
      if (rows.length >= LIMIT) throw new Error(`more than ${LIMIT} posts in range - ask for a shorter range`);
      return rows.map(flat);
    },
    async getPost(id) {
      const r = must(await sb.from("social_posts").select(COLS).eq("id", id).maybeSingle());
      return r ? flat(r) : null;
    },
    async findByTitleBetween(title, fromIso, toIso) {
      const rows = must(await sb.from("social_posts").select(COLS).is("deleted_at", null)
        .ilike("title", title.replace(/[%_\\]/g, (c) => "\\" + c))
        .gte("scheduled_for", fromIso).lt("scheduled_for", toIso)) as any[];
      return rows.map(flat);
    },
    async insertPost(row) {
      return flat(must(await sb.from("social_posts").insert(row).select(COLS).single()));
    },
    async updatePostIfStage(id, expectStage, patch) {
      const r = must(await sb.from("social_posts").update(patch).eq("id", id).eq("stage", expectStage)
        .is("deleted_at", null).select(COLS).maybeSingle());
      return r ? flat(r) : null;
    },
    async insertNote(postId, body, authorName) {
      return must(await sb.from("social_post_notes")
        .insert({ post_id: postId, body, author_name: authorName, author_id: null })
        .select("id,created_at").single()) as { id: string; created_at: string };
    },
    async profileIdByEmail(email) {
      const r = must(await sb.from("profiles").select("id").eq("email", email).is("deleted_at", null).maybeSingle()) as any;
      return r ? r.id : null;
    },
    async log(entry) {
      await sb.from("connector_log").insert({ connector: "social-mcp", ...entry });
    },
  };
}
