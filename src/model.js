// What the document should look like. Works on a plain snapshot of the
// document, so everything here runs without Word.
//
// The names below are the stored format. The README documents them, and no
// other file spells them out.

/** Namespace and root element of the custom XML part that holds the store. */
export const NS = "urn:siglum:store:1";
/** Namespace of the custom XML part that holds the last saved store. */
export const BACKUP_NS = "urn:siglum:store:backup:1";
export const ROOT = "siglum";
/** Tag of a citation's inline content control, followed by the citation's id. */
export const CITE_PREFIX = "siglum:cite:";
/** Tag of a bibliography's block content control. */
export const BIB_TAG = "siglum:bibliography";
/** A citation's title names its keys after this prefix: `Cite: key1; key2`. */
export const TITLE_PREFIX = "Cite: ";
/** Text of a new citation until it is numbered. */
export const CITE_NEW = "[?]";
/** Text of a bibliography without citations, unless the caller passes a translation. */
export const BIB_EMPTY = "No sources cited.";

import { runsText } from "./render.js";
import { KEY } from "./library.js";

/**
 * The store: `sources` maps a key to a CSL-JSON item, `citations` maps a
 * citation's id to its items [{id, locator?, label?, prefix?, suffix?,
 * suppressAuthor?}], and `settings`
 * holds what belongs to the document: `locale`, the language of its citations
 * and bibliographies, and `style`, the id of its citation style.
 */
export function emptyStore(locale) {
  return { sources: {}, citations: {}, settings: { locale, style: "nlm" } };
}

/** The bundled citation styles, by id. `nlm` is the default. */
export const STYLES = ["nlm", "apa"];

/** The store's style if it names a bundled one, else the default. */
export function resolveStyle(settings) {
  const style = settings?.style;
  return STYLES.includes(style) ? style : "nlm";
}

const escapeXml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const unescapeXml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const PAYLOAD = new RegExp(`<${ROOT}\\b[^>]*>([\\s\\S]*)</${ROOT}>`);

export function storeToXml({ sources, citations, settings }, ns = NS) {
  return `<?xml version="1.0" encoding="UTF-8"?><${ROOT} xmlns="${ns}">${escapeXml(JSON.stringify({ sources, citations, settings }))}</${ROOT}>`;
}

const isRecord = (x) => x !== null && typeof x === "object" && !Array.isArray(x);

export function storeFromXml(xml) {
  const m = PAYLOAD.exec(xml);
  if (!m) throw new Error("The document's Siglum store is damaged.");
  let data;
  try { data = JSON.parse(unescapeXml(m[1])); } catch { data = null; }
  const { sources, citations, settings } = data ?? {};
  if (!isRecord(sources) || !isRecord(citations) || !isRecord(settings) || typeof settings.locale !== "string") {
    throw new Error("The document's Siglum store is damaged.");
  }
  return { sources, citations, settings };
}

export function newId() {
  const a = new Uint8Array(6);
  globalThis.crypto.getRandomValues(a);
  return Array.from(a, (b) => (b % 36).toString(36)).join("") + Date.now().toString(36).slice(-4);
}

export function citeTitle(items) {
  const t = TITLE_PREFIX + items.map((x) => x.id).join("; ");
  return t.length <= 64 ? t : t.slice(0, 63) + "…";
}

/** Keys from a title, unless it was cut short. */
export function keysFromTitle(title) {
  if (!title?.startsWith(TITLE_PREFIX) || title.endsWith("…")) return null;
  const keys = title.slice(TITLE_PREFIX.length).split(";").map((k) => k.trim()).filter(Boolean);
  return keys.length ? keys : null;
}

// ---- typed markers ------------------------------------------------------------
//
// The syntax is Typst's: `@key` cites a source, `@key[p. 12]` adds a
// supplement, and markers with only spaces between them are one citation.

// What may come right before the "@": the start of the paragraph, white
// space, an opening parenthesis, bracket or brace, a quotation mark, a dash, a
// comma or a semicolon. So "name@example.org" and "example.org/@name" are text.
const BEFORE = /[\s(\[\{"'“”„‘’‚«»‹›–—,;]/u;
// The "@" and its key, which is not cut short: `@a@b` is no marker, and
// neither is a shortened key. A combining mark belongs to the key it follows,
// and a key never ends before a joiner, so a marker never ends inside a
// character as the reader sees it.
const MARKER = new RegExp(String.raw`@(${KEY})(?![\p{L}\p{M}\p{N}_\-@\u200c\u200d])`, "uy");
// Only spaces and no-break spaces join markers.
const GAP = /[  ]+/y;
// A supplement holds no tab, line break or other control character.
const CONTROL = /\p{Cc}/u;
const ATTACHES = /[\p{M}\u200c\u200d]/u;

/**
 * The end of the balanced brackets that open at `i`, or -1. Brackets inside
 * count in pairs, as in Typst: `[outer[inner]]` is one supplement.
 */
function closing(text, i) {
  let depth = 0;
  for (let j = i; j < text.length; j++) {
    const c = text[j];
    if (CONTROL.test(c)) return -1;
    if (c === "[") depth++;
    else if (c === "]" && --depth === 0) return ATTACHES.test(text[j + 1] ?? "") ? -1 : j + 1;
  }
  return -1;
}

function markerAt(text, i) {
  MARKER.lastIndex = i;
  const m = MARKER.exec(text);
  if (!m) return null;
  const id = m[1].normalize("NFC");
  const end = text[MARKER.lastIndex] === "[" ? closing(text, MARKER.lastIndex) : -1;
  if (end === -1) return { item: { id }, end: MARKER.lastIndex };
  return { item: { id, ...parseSupplement(text.slice(MARKER.lastIndex + 1, end - 1)) }, end };
}

/**
 * Whether a content control that reads `control` can share a character with a
 * marker group that reads `group`: one holds the other, or one ends as the
 * other begins. Only then is it worth asking Word where the two ranges are.
 */
export function mayTouch(group, control) {
  if (group.includes(control) || control.includes(group)) return true;
  const n = Math.min(group.length, control.length);
  for (let k = 1; k < n; k++) {
    if (control.endsWith(group.slice(0, k)) || group.endsWith(control.slice(0, k))) return true;
  }
  return false;
}

/**
 * The markers in one paragraph's text, adjacent ones merged into groups:
 * [{start, end, text, items}]. `text` is the group as typed, and the keys in
 * `items` are in Unicode's composed form (NFC), as the library stores them.
 */
export function findMarkers(text) {
  const groups = [];
  let i = 0;
  while ((i = text.indexOf("@", i)) !== -1) {
    const first = i === 0 || BEFORE.test(text[i - 1]) ? markerAt(text, i) : null;
    if (!first) { i++; continue; }
    const items = [first.item];
    const seen = new Set([JSON.stringify(first.item)]);
    let end = first.end;
    for (;;) {
      GAP.lastIndex = end;
      const next = GAP.test(text) ? markerAt(text, GAP.lastIndex) : null;
      if (!next) break;
      // A source cited twice in a row, without a supplement, counts once.
      const same = JSON.stringify(next.item);
      if (!seen.has(same)) { seen.add(same); items.push(next.item); }
      end = next.end;
    }
    groups.push({ start: i, end, text: text.slice(i, end), items });
    i = end;
  }
  return groups;
}

// CSL locator labels and the abbreviations people type for them, English and German.
export const LABELS = {
  page: ["p", "pp", "pg", "page", "pages", "s", "seite", "seiten"],
  chapter: ["ch", "chap", "chapter", "chapters", "kap", "kapitel"],
  section: ["sec", "sect", "section", "sections", "§", "§§", "abschn", "abschnitt", "abschnitte"],
  paragraph: ["para", "paras", "paragraph", "paragraphs", "¶", "abs", "absatz"],
  figure: ["fig", "figs", "figure", "figures", "abb", "abbildung"],
  volume: ["vol", "vols", "volume", "volumes", "bd", "band"],
  line: ["l", "ll", "line", "lines", "z", "zeile", "zeilen"],
  note: ["n", "nn", "note", "notes", "anm", "anmerkung"],
  column: ["col", "cols", "column", "columns", "sp", "spalte", "spalten"],
};
const LABEL_OF = new Map(Object.entries(LABELS).flatMap(([label, words]) => words.map((w) => [w, label])));
const ARABIC = String.raw`\d+[a-z]?`;
const ROMAN = String.raw`(?=[ivxlcdm])m*(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{0,3})`;
const numbers = (n) => new RegExp(String.raw`^(?:${n})(?:\s*[-–—,&]\s*(?:${n}))*(?:\s?ff?\.?)?$`, "i");
const BARE = numbers(ARABIC);
const LABELLED = numbers(`${ARABIC}|${ROMAN}`);

/**
 * Reads a supplement, the text in `@key[…]`. A locator such as "p. 12",
 * "pp. 12-15", "ch. 3", "S. 12" or "12" comes back as {locator, label} for the
 * CSL processor. Any other text comes back as {suffix}, and an empty one as {}.
 */
export function parseSupplement(text) {
  const t = String(text ?? "").replace(/\u00a0/g, " ").trim();
  if (!t) return {};
  const dashes = (s) => s.replace(/\s*[–—-]\s*/g, "-");
  if (BARE.test(t)) return { locator: dashes(t), label: "page" };
  const m = /^([\p{L}§¶]+)\.?\s*(.+)$/u.exec(t);
  const label = m && LABEL_OF.get(m[1].toLowerCase());
  if (label && LABELLED.test(m[2])) return { locator: dashes(m[2]), label };
  return { suffix: t };
}

// ---- planning an update -------------------------------------------------------

const norm = (s) => String(s ?? "").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ").trim();
const isEmpty = (s) => norm(s) === "";

/**
 * Decides which citations each bibliography lists: [{citations, bibs}], each
 * numbered on its own. A bibliography takes the citations since the one before
 * it, and citations after the last one belong to it, for example in
 * appendices. This is the one place that knows the rule, so a document
 * setting for a single shared bibliography only has to change this function.
 */
function sections(sequence) {
  const out = [{ citations: [], bibs: [] }];
  for (const x of sequence) {
    if (x.bib == null) { out[out.length - 1].citations.push(x); continue; }
    out[out.length - 1].bibs.push(x.bib);
    out.push({ citations: [], bibs: [] });
  }
  const tail = out.pop();
  if (out.length) out[out.length - 1].citations.push(...tail.citations);
  else out.push(tail);
  return out;
}

/** A bibliography's paragraphs, from its text as Word gives it: paragraphs end in "\r". */
export const paragraphsOf = (text) => {
  const ps = String(text ?? "").split("\r");
  if (ps.length > 1 && ps[ps.length - 1] === "") ps.pop();
  return ps;
};

// Beyond this many changed entries, or this share of them, a bibliography is
// written whole: Word takes a whole bibliography in one piece faster than
// many single entries.
const EDITS_MAX = 24;
const EDITS_SHARE = 0.25;

/**
 * How to turn the paragraphs `have` into the entries `want`, both as texts:
 * hunks {at, remove, insert}, where `remove` paragraphs from index `at` of
 * `have` give way to the entries `insert` (indices into `want`). The hunks
 * are in document order and do not overlap. Returns null when so much
 * differs that writing the whole bibliography is cheaper.
 */
export function diffEntries(have, want) {
  let start = 0;
  while (start < have.length && start < want.length && have[start] === want[start]) start++;
  let end = 0;
  while (end < have.length - start && end < want.length - start && have[have.length - 1 - end] === want[want.length - 1 - end]) end++;
  const a = have.slice(start, have.length - end);
  const b = want.slice(start, want.length - end);
  const limit = Math.max(EDITS_MAX, Math.ceil(want.length * EDITS_SHARE));
  if (Math.min(a.length, b.length) > limit || a.length * b.length > 4e6) return null;
  // Longest common subsequence of the middle parts.
  const n = a.length;
  const m = b.length;
  const len = new Uint32Array((n + 1) * (m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      len[i * (m + 1) + j] = a[i] === b[j] ? len[(i + 1) * (m + 1) + j + 1] + 1 : Math.max(len[(i + 1) * (m + 1) + j], len[i * (m + 1) + j + 1]);
    }
  }
  const hunks = [];
  let hunk = null;
  let changed = 0;
  const flush = () => { if (hunk) { hunks.push(hunk); changed += Math.max(hunk.remove, hunk.insert.length); hunk = null; } };
  let i = 0;
  let j = 0;
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) { flush(); i++; j++; continue; }
    hunk ??= { at: start + i, remove: 0, insert: [] };
    if (j < m && (i === n || len[i * (m + 1) + j + 1] >= len[(i + 1) * (m + 1) + j])) hunk.insert.push(start + j++);
    else { hunk.remove++; i++; }
  }
  flush();
  return changed > limit ? null : hunks;
}

/**
 * Plans an update. `controls` are the citations and bibliographies in document
 * order: {kind: "cite", tag, title, current, original} or {kind: "bib", current}.
 * `current` and `original` are the texts with tracked changes accepted and
 * rejected, "" while Word shows the placeholder; a bibliography's paragraphs
 * end in "\r". `locale` is the language of citations and bibliographies,
 * `bibEmpty` the text of a bibliography without citations, and `elsewhere`
 * the ids of citations outside the controls, such as in footnotes, whose
 * stored items stay. Returns what to write, by position in `controls`:
 *
 * - `citations`: [{index, runs}], the citations whose text differs;
 * - `bibs`: [{index, entries, hunks}], the bibliographies that differ, with
 *   `hunks` from diffEntries, or null to write them whole;
 * - `removes`: controls emptied by hand;
 * - `store`: the store to save, when it changed: citations pasted from
 *   another document join it, and citations that are gone leave it.
 *
 * It depends on nothing but its arguments.
 */
export function plan(controls, store, renderer, { locale = "en-US", bibEmpty = BIB_EMPTY, elsewhere = [] } = {}) {
  const sequence = [];
  const out = { citations: [], bibs: [], removes: [], unknown: [], pending: 0, store: null };
  const citations = {};
  const kept = new Set(elsewhere);
  let changed = false;
  controls.forEach((c, index) => {
    if (c.kind === "bib") { sequence.push({ bib: index }); return; }
    const id = c.tag.slice(CITE_PREFIX.length);
    // Emptied by hand: the control is a leftover. A tracked deletion waits for review.
    if (isEmpty(c.current)) {
      if (isEmpty(c.original)) { out.removes.push(index); return; }
      out.pending++;
      kept.add(id);
      return;
    }
    kept.add(id);
    let items = citations[id] ?? store.citations[id];
    if (!items) {
      // Pasted from another document: the title still names the keys.
      const keys = keysFromTitle(c.title);
      if (keys && keys.every((k) => store.sources[k])) {
        items = keys.map((k) => ({ id: k }));
        changed = true;
      }
    }
    if (items) citations[id] = items;
    const missing = items ? items.filter((x) => !store.sources[x.id]).map((x) => x.id) : keysFromTitle(c.title) ?? [];
    if (!items || missing.length) {
      out.unknown.push({ index, keys: missing, text: norm(c.current) });
      return;
    }
    sequence.push({ index, items, current: c.current });
  });
  // Stored citations of controls that waited for review or are elsewhere stay.
  for (const id of kept) if (!citations[id] && store.citations[id]) citations[id] = store.citations[id];
  if (changed || Object.keys(citations).length !== Object.keys(store.citations).length) out.store = { ...store, citations };

  const secs = sections(sequence);
  const rendered = renderer.render(secs, store.sources, locale);
  secs.forEach((sec, s) => {
    sec.citations.forEach((c, i) => {
      const runs = rendered[s].citations[i];
      if (norm(runsText(runs)) !== norm(c.current)) out.citations.push({ index: c.index, runs });
    });
    const entries = rendered[s].entries.length ? rendered[s].entries : [[{ text: bibEmpty }]];
    const want = entries.map((e) => norm(runsText(e)));
    for (const bib of sec.bibs) {
      const have = paragraphsOf(controls[bib].current).map(norm);
      if (want.length === have.length && want.every((w, k) => w === have[k])) continue;
      out.bibs.push({ index: bib, entries, hunks: diffEntries(have, want) });
    }
  });
  out.sections = secs.length;
  out.cited = secs.reduce((n, s) => n + s.citations.length, 0);
  return out;
}
