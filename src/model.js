// What the document should look like. Works on a plain snapshot of the
// document, so everything here runs without Word.
//
// The names below are the stored format. The README documents them, and no
// other file spells them out.

/** Namespace and root element of the custom XML part that holds the store. */
export const NS = "urn:siglum:store:1";
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
 * citation's id to its items [{id, locator?, label?, suffix?}], and `settings`
 * holds what belongs to the document: `locale`, the language of its citations
 * and bibliographies.
 */
export function emptyStore(locale) {
  return { sources: {}, citations: {}, settings: { locale } };
}

const escapeXml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const unescapeXml = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const PAYLOAD = new RegExp(`<${ROOT}\\b[^>]*>([\\s\\S]*)</${ROOT}>`);

export function storeToXml({ sources, citations, settings }) {
  return `<?xml version="1.0" encoding="UTF-8"?><${ROOT} xmlns="${NS}">${escapeXml(JSON.stringify({ sources, citations, settings }))}</${ROOT}>`;
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
const LABELS = {
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
const lines = (s) => String(s ?? "").split(/[\r\n\v]+/).map(norm).filter(Boolean);
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

/**
 * Plans an update. `controls` are the citations and bibliographies in document
 * order: {kind: "cite", tag, title, current, original} or {kind: "bib", current}.
 * `current` and `original` are the texts with tracked changes accepted and
 * rejected, "" while Word shows the placeholder. `locale` is the language of
 * citations and bibliographies, and `bibEmpty` the text of a bibliography
 * without citations. Returns what to write, by position in `controls`.
 */
export function plan(controls, store, renderer, { locale = "en-US", bibEmpty = BIB_EMPTY } = {}) {
  const sequence = [];
  const out = { citations: [], bibs: [], removes: [], unknown: [], pending: 0, store: null };
  let citations = store.citations;
  controls.forEach((c, index) => {
    if (c.kind === "bib") { sequence.push({ bib: index }); return; }
    // Emptied by hand: the control is a leftover. A tracked deletion waits for review.
    if (isEmpty(c.current)) {
      if (isEmpty(c.original)) out.removes.push(index);
      else out.pending++;
      return;
    }
    const id = c.tag.slice(CITE_PREFIX.length);
    let items = citations[id];
    if (!items) {
      // Pasted from another document: the title still names the keys.
      const keys = keysFromTitle(c.title);
      if (keys && keys.every((k) => store.sources[k])) {
        items = keys.map((k) => ({ id: k }));
        citations = { ...citations, [id]: items };
      }
    }
    const missing = items ? items.filter((x) => !store.sources[x.id]).map((x) => x.id) : keysFromTitle(c.title) ?? [];
    if (!items || missing.length) {
      out.unknown.push({ index, keys: missing, text: norm(c.current) });
      return;
    }
    sequence.push({ index, items, current: c.current });
  });
  if (citations !== store.citations) out.store = { ...store, citations };

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
      const have = lines(controls[bib].current);
      if (want.join("\n") !== have.join("\n")) out.bibs.push({ index: bib, entries });
    }
  });
  out.sections = secs.length;
  out.cited = secs.reduce((n, s) => n + s.citations.length, 0);
  return out;
}
