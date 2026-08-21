// pdf.ts — a small, dependency-free PDF writer.
//
// WHY THIS EXISTS. An invoice has to arrive as a real .pdf: it is the artifact a
// client forwards to their bookkeeper and files for their own accounts. The repo
// had no PDF path — agreements and the social snapshot both produce HTML and lean
// on the browser's Save-as-PDF, which is fine for something you look at and wrong
// for something you send.
//
// Pulling in a PDF library would mean an npm dependency inside an edge function
// that gets deployed by a script and cannot be run locally. This module is the
// alternative: ~300 lines, no imports, and unit-testable from Node (which is why
// it is a separate .ts file, the same reason scan-gear-market/scoring.ts is).
//
// SCOPE, deliberately narrow. Standard-14 Type1 fonts (Helvetica and
// Helvetica-Bold), text, lines and filled rectangles. That is everything an
// invoice needs and nothing else. No images, no embedded fonts, no compression —
// an uncompressed content stream is a few KB larger and enormously easier to
// prove correct, and you can read the output in a text editor when something
// looks wrong.
//
// TWO THINGS THAT MUST STAY RIGHT, because a PDF that violates either will open
// in one viewer and fail in another:
//   1. The xref table holds the BYTE OFFSET of every object. Everything is
//      assembled as bytes and measured as bytes for that reason — never as a
//      string whose length in characters is not its length in bytes.
//   2. Text is written in WinAnsi (Latin-1), one byte per glyph. Anything
//      outside that range is transliterated, not dropped silently mid-word.

const LETTER_W = 612;
const LETTER_H = 792;

// Adobe standard AFM advance widths, 1/1000 em, for code points 32..126.
// Used for right-alignment and for wrapping. A wrong number here misaligns a
// column; it cannot produce an invalid file.
const W_REG = [
  278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
  556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
];
const W_BOLD = [
  278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
  975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
  333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
  611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
];

// Characters that routinely arrive from a database of artist and venue names and
// have no WinAnsi code point. Transliterating beats printing a black diamond in
// the middle of a client's company name.
const TRANSLIT: Record<string, string> = {
  "‘": "'", "’": "'", "‚": ",", "“": '"', "”": '"',
  "–": "-", "—": "-", "…": "...", "•": "-", " ": " ",
  "−": "-", "×": "x", "⁄": "/", "™": "(TM)", "€": "EUR",
  "ł": "l", "ø": "o", "Ø": "O", "đ": "d", "Đ": "D",
};

/** Fold text to something WinAnsi can actually represent. */
export function toWinAnsi(input: string): string {
  let s = String(input ?? "");
  for (const [from, to] of Object.entries(TRANSLIT)) s = s.split(from).join(to);
  // Decompose accents (é -> e + combining) and drop the combining marks, so a
  // name keeps its letters even when the glyph is unavailable.
  let out = "";
  for (const ch of s.normalize("NFC")) {
    const c = ch.codePointAt(0)!;
    if (c === 9) { out += "    "; continue; }
    if (c === 10 || c === 13) { out += ch; continue; }
    if (c < 32) continue;
    if (c <= 255) { out += ch; continue; }
    const folded = ch.normalize("NFD").replace(/[̀-ͯ]/g, "");
    out += folded.codePointAt(0)! <= 255 ? folded : "?";
  }
  return out;
}

/** Width of a string in points at a given size. */
export function measure(text: string, size: number, bold = false): number {
  const table = bold ? W_BOLD : W_REG;
  let w = 0;
  const s = toWinAnsi(text);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    // Everything outside the metric table is charged at the width of a lowercase
    // n — close enough for accented Latin, and never zero, which would let a
    // long string report as fitting when it does not.
    w += c >= 32 && c <= 126 ? table[c - 32] : (bold ? 611 : 556);
  }
  return (w * size) / 1000;
}

/** Greedy word wrap to a pixel width. Returns at least one line. */
export function wrap(text: string, size: number, maxWidth: number, bold = false): string[] {
  const paras = toWinAnsi(text).split(/\r?\n/);
  const out: string[] = [];
  for (const para of paras) {
    const words = para.split(/\s+/).filter((w) => w.length);
    if (!words.length) { out.push(""); continue; }
    let line = "";
    for (const word of words) {
      const probe = line ? line + " " + word : word;
      if (measure(probe, size, bold) <= maxWidth || !line) {
        // A single word longer than the column is hard-split rather than allowed
        // to run off the page edge.
        if (!line && measure(word, size, bold) > maxWidth) {
          let chunk = "";
          for (const ch of word) {
            if (measure(chunk + ch, size, bold) > maxWidth && chunk) { out.push(chunk); chunk = ch; }
            else chunk += ch;
          }
          line = chunk;
        } else line = probe;
      } else { out.push(line); line = word; }
    }
    if (line) out.push(line);
  }
  return out.length ? out : [""];
}

type RGB = [number, number, number];
const HEX = (hex: string): RGB => {
  const h = hex.replace("#", "");
  return [
    parseInt(h.slice(0, 2), 16) / 255,
    parseInt(h.slice(2, 4), 16) / 255,
    parseInt(h.slice(4, 6), 16) / 255,
  ];
};
const n3 = (v: number) => (Math.round(v * 1000) / 1000).toString();

export type TextOpts = {
  size?: number;
  bold?: boolean;
  color?: string;
  align?: "left" | "right" | "center";
  width?: number;      // required for right/center alignment
  spacing?: number;    // letter spacing, points
};

/**
 * A page in a top-left coordinate system: y grows DOWNWARD from the top edge,
 * because every layout decision in an invoice is "how far down the page am I".
 * The conversion to PDF's bottom-left origin happens in one place, on write.
 */
class Page {
  ops: string[] = [];
  width: number;
  height: number;
  // Written out rather than as constructor parameter properties: Node's
  // strip-only TypeScript mode rejects those, and this file has to stay
  // importable from a plain `node scripts/...` test.
  constructor(width: number, height: number) {
    this.width = width;
    this.height = height;
  }

  private esc(s: string) {
    return s.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
  }

  text(x: number, y: number, str: string, o: TextOpts = {}) {
    const size = o.size ?? 10;
    const bold = !!o.bold;
    const s = toWinAnsi(str);
    if (!s) return;
    let tx = x;
    if (o.align === "right" || o.align === "center") {
      const w = measure(s, size, bold) + (o.spacing ? o.spacing * (s.length - 1) : 0);
      const box = o.width ?? 0;
      tx = o.align === "right" ? x + box - w : x + (box - w) / 2;
    }
    const [r, g, b] = HEX(o.color ?? "#1A1410");
    this.ops.push(
      `BT /${bold ? "F2" : "F1"} ${n3(size)} Tf ${n3(r)} ${n3(g)} ${n3(b)} rg` +
        (o.spacing ? ` ${n3(o.spacing)} Tc` : "") +
        ` 1 0 0 1 ${n3(tx)} ${n3(this.height - y)} Tm (${this.esc(s)}) Tj` +
        (o.spacing ? " 0 Tc" : "") + " ET",
    );
  }

  line(x1: number, y1: number, x2: number, y2: number, o: { w?: number; color?: string } = {}) {
    const [r, g, b] = HEX(o.color ?? "#1A1410");
    this.ops.push(
      `${n3(r)} ${n3(g)} ${n3(b)} RG ${n3(o.w ?? 0.75)} w ` +
        `${n3(x1)} ${n3(this.height - y1)} m ${n3(x2)} ${n3(this.height - y2)} l S`,
    );
  }

  rect(x: number, y: number, w: number, h: number, o: { fill?: string } = {}) {
    const [r, g, b] = HEX(o.fill ?? "#EEEEEE");
    this.ops.push(
      `${n3(r)} ${n3(g)} ${n3(b)} rg ${n3(x)} ${n3(this.height - y - h)} ${n3(w)} ${n3(h)} re f`,
    );
  }

  content(): string { return this.ops.join("\n"); }
}

export class Pdf {
  pages: Page[] = [];
  title: string;
  author: string;
  constructor(title = "Invoice", author = "Come With") {
    this.title = title;
    this.author = author;
  }

  addPage(width = LETTER_W, height = LETTER_H): Page {
    const p = new Page(width, height);
    this.pages.push(p);
    return p;
  }

  /** Assemble the file. Offsets are measured in BYTES, not characters. */
  build(): Uint8Array {
    const chunks: number[] = [];
    const push = (s: string) => { for (let i = 0; i < s.length; i++) chunks.push(s.charCodeAt(i) & 0xff); };
    const at = () => chunks.length;

    // Object numbering: 1 catalog, 2 pages, 3 F1, 4 F2, 5 info,
    // then per page a page object and a content object.
    const nPages = this.pages.length;
    const firstPageObj = 6;
    const offsets: number[] = [];
    const obj = (num: number, body: string) => {
      offsets[num] = at();
      push(`${num} 0 obj\n${body}\nendobj\n`);
    };

    push("%PDF-1.4\n");
    // A binary comment marks the file as binary for tools that sniff it.
    chunks.push(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a);

    const kids = this.pages.map((_, i) => `${firstPageObj + i * 2} 0 R`).join(" ");
    obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
    obj(2, `<< /Type /Pages /Kids [${kids}] /Count ${nPages} >>`);
    obj(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
    obj(4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");
    const clean = (s: string) => toWinAnsi(s).replace(/[()\\]/g, "");
    obj(5, `<< /Title (${clean(this.title)}) /Author (${clean(this.author)}) /Producer (Come With dashboard) >>`);

    this.pages.forEach((p, i) => {
      const pageNum = firstPageObj + i * 2;
      const contentNum = pageNum + 1;
      obj(
        pageNum,
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${n3(p.width)} ${n3(p.height)}] ` +
          `/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${contentNum} 0 R >>`,
      );
      const body = p.content();
      // /Length is the BYTE length. Every char is written as one masked byte
      // (push() does `& 0xff`) and toWinAnsi() has already guaranteed nothing
      // above U+00FF reaches here, so char length and byte length are the same.
      const bytes = body.length;
      offsets[contentNum] = at();
      push(`${contentNum} 0 obj\n<< /Length ${bytes} >>\nstream\n`);
      push(body);
      push("\nendstream\nendobj\n");
    });

    const maxObj = firstPageObj + nPages * 2 - 1;
    const xrefAt = at();
    push(`xref\n0 ${maxObj + 1}\n`);
    push("0000000000 65535 f \n");
    for (let i = 1; i <= maxObj; i++) {
      push(String(offsets[i] ?? 0).padStart(10, "0") + " 00000 n \n");
    }
    push(`trailer\n<< /Size ${maxObj + 1} /Root 1 0 R /Info 5 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`);

    return new Uint8Array(chunks);
  }
}

export const PAGE = { W: LETTER_W, H: LETTER_H };
