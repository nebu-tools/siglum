// Formats citations and bibliographies with citeproc-js. A run is
// {text, italic?, bold?, superscript?, subscript?, smallCaps?}.

import CSL from "citeproc";

/** Turns citeproc's HTML into runs. Block elements become paragraph breaks ("\n"). */
export function htmlToRuns(html) {
  const runs = [];
  const stack = [{}];
  const re = /<(\/?)([a-z]+)([^>]*)>|([^<]+)/gi;
  let m;
  while ((m = re.exec(html))) {
    if (m[4] != null) {
      const text = decode(m[4]);
      if (!text) continue;
      const fmt = stack[stack.length - 1];
      const prev = runs[runs.length - 1];
      if (prev && sameFormat(prev, fmt)) prev.text += text;
      else runs.push({ text, ...fmt });
      continue;
    }
    const [, close, tag, attrs] = m;
    if (close) { if (stack.length > 1) stack.pop(); continue; }
    if (tag === "br") { runs.push({ text: "\n" }); continue; }
    const f = { ...stack[stack.length - 1] };
    if (tag === "i" || tag === "em") f.italic = !f.italic;
    if (tag === "b" || tag === "strong") f.bold = true;
    if (tag === "sup") f.superscript = true;
    if (tag === "sub") f.subscript = true;
    if (/font-style:\s*italic/.test(attrs)) f.italic = true;
    if (/font-style:\s*normal/.test(attrs)) f.italic = false;
    if (/font-weight:\s*bold/.test(attrs)) f.bold = true;
    if (/font-variant:\s*small-caps/.test(attrs)) f.smallCaps = true;
    if (/vertical-align:\s*sup/.test(attrs)) f.superscript = true;
    if (/vertical-align:\s*sub/.test(attrs)) f.subscript = true;
    for (const k of Object.keys(f)) if (!f[k]) delete f[k];
    stack.push(f);
  }
  return runs;
}

const FLAGS = ["italic", "bold", "superscript", "subscript", "smallCaps"];
const sameFormat = (a, b) => FLAGS.every((k) => !!a[k] === !!b[k]);

function decode(s) {
  return s
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&amp;/g, "&")
    .replace(/\s*\n\s*/g, "");
}

export const runsText = (runs) => runs.map((r) => r.text).join("");

/**
 * One bibliography entry: number, tab, text. citeproc puts the number in
 * csl-left-margin when the style aligns it; other styles give plain text.
 */
function entryRuns(html) {
  const left = /<div class="csl-left-margin">([\s\S]*?)<\/div>/.exec(html);
  const right = /<div class="csl-right-inline">([\s\S]*?)<\/div>/.exec(html);
  if (left && right) return [...htmlToRuns(left[1]), { text: "\t" }, ...htmlToRuns(right[1])];
  return htmlToRuns(html.replace(/<\/?div[^>]*>/g, ""));
}

/** The bundled locale for a language tag: the tag itself, one of the same language, or en-US. */
export function resolveLocale(tag, available) {
  const want = String(tag ?? "").replace("_", "-").toLowerCase();
  const lang = want.split("-")[0];
  return available.find((l) => l.toLowerCase() === want)
    ?? available.find((l) => l.toLowerCase().split("-")[0] === lang)
    ?? (available.includes("en-US") ? "en-US" : available[0]);
}

/**
 * Creates a renderer for one style. `locales` maps a language tag to locale XML.
 * render(sections, library, lang) takes [{citations: [{items: [{id, locator?, label?, suffix?}]}]}]
 * and returns, per section, the runs of each citation and of each entry, in
 * the bundled locale that fits `lang`.
 *
 * Each section is numbered by a fresh engine that learns the sources in the
 * order of their first citation and then formats one citation at a time,
 * which is many times faster than citeproc's rebuild of a whole document. It
 * keeps no position between citations, so a style's "ibid." does not occur;
 * the document's citations are in the text, not in notes. A section that is
 * the same as before (the same citations, the same sources, the same language)
 * comes from a cache, so a refresh where little changed formats little.
 */
export function createRenderer({ style, locales, cacheSize = 64 }) {
  const available = Object.keys(locales);
  const resolve = (tag) => resolveLocale(tag, available);
  const cache = new Map();
  const remember = (key, value) => {
    cache.delete(key);
    cache.set(key, value);
    if (cache.size > cacheSize) cache.delete(cache.keys().next().value);
    return value;
  };
  function section(sec, library, locale) {
    const order = [];
    const seen = new Set();
    for (const c of sec.citations) for (const x of c.items) if (!seen.has(x.id)) { seen.add(x.id); order.push(x.id); }
    const key = JSON.stringify([locale, sec.citations.map((c) => c.items), order.map((id) => library[id])]);
    const hit = cache.get(key);
    if (hit) return remember(key, hit);
    const engine = new CSL.Engine({ retrieveLocale: (l) => locales[resolve(l)], retrieveItem: (id) => library[id] }, style, locale, true);
    engine.updateItems(order);
    const citations = sec.citations.map((c) => htmlToRuns(engine.makeCitationCluster(c.items.map(citationItem))));
    const [, entries] = engine.makeBibliography();
    return remember(key, { citations, entries: entries.map(entryRuns) });
  }
  return {
    locales: available,
    resolve,
    render(sections, library, lang = "en-US") {
      const locale = resolve(lang);
      return sections.map((sec) => (sec.citations.length ? section(sec, library, locale) : { citations: [], entries: [] }));
    },
  };
}

function citationItem(x) {
  const item = { id: x.id };
  if (x.locator) { item.locator = x.locator; item.label = x.label; }
  if (x.suffix) item.suffix = ", " + x.suffix;
  return item;
}
