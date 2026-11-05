// Render harness only: answers control.html's selects from a prod snapshot.
const SNAP = await (await fetch('./snap.json')).json();
class Q {
  constructor(t) { this.t = t; this.f = []; this.mode = null; this.rng = null; this.write = false; }
  select() { return this; } order() { return this; } limit() { return this; } or() { return this; }
  eq(c, v) { this.f.push(r => String(r[c]) === String(v)); return this; }
  is(c, v) { this.f.push(r => (r[c] ?? null) === v); return this; }
  in(c, vs) { const s = new Set(vs.map(String)); this.f.push(r => s.has(String(r[c]))); return this; }
  ilike() { return this; }
  range(a, b) { this.rng = [a, b]; return this; }
  single() { this.mode = 'one'; return this; } maybeSingle() { this.mode = 'one'; return this; }
  insert() { this.write = true; return this; } update() { this.write = true; return this; } delete() { this.write = true; return this; } upsert() { this.write = true; return this; }
  then(res, rej) {
    if (this.write) return Promise.resolve({ data: [], error: null }).then(res, rej);
    let rows = (SNAP[this.t] || []).filter(r => this.f.every(fn => fn(r)));
    if (this.rng) rows = rows.slice(this.rng[0], this.rng[1] + 1);
    return Promise.resolve({ data: this.mode === 'one' ? (rows[0] ?? null) : rows, error: null }).then(res, rej);
  }
}
export function createClient() {
  return {
    from: (t) => new Q(t),
    auth: { getSession: async () => ({ data: { session: { user: { id: 'u1' } } } }), onAuthStateChange: () => {} },
  };
}
