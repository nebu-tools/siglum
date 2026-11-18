// Reads and writes the document through Office.js. All decisions are in model.js.
//
// Every operation takes a session: {renderer, locale}. `locale` is the
// language a new document gets for its citations and bibliographies.
/* global Word */

import {
  NS, CITE_PREFIX, BIB_TAG, CITE_NEW,
  emptyStore, storeToXml, storeFromXml, newId, citeTitle, findMarkers, plan,
} from "./model.js";
import { documentText, UserError } from "./messages.js";

// Hanging indent of bibliography entries, in points (1 cm).
const HANG = 28.35;

const isCite = (cc) => cc.tag?.startsWith(CITE_PREFIX);
const isOurs = (cc) => isCite(cc) || cc.tag === BIB_TAG;

/**
 * The content controls that can hold citations and bibliographies. Citations
 * in footnotes and endnotes would be added here.
 */
function citationControls(context) {
  return context.document.body.contentControls;
}

async function saveStore(context, store) {
  const parts = context.document.customXmlParts.getByNamespace(NS);
  parts.load("items");
  await context.sync();
  if (parts.items.length) parts.items[0].setXml(storeToXml(store));
  else context.document.customXmlParts.add(storeToXml(store));
  await context.sync();
}

/**
 * The document's store. Every operation starts here. A document without one
 * gets an empty store in the session's language; the next save keeps it.
 */
async function openStore(context, session) {
  const parts = context.document.customXmlParts.getByNamespace(NS);
  parts.load("items");
  await context.sync();
  if (!parts.items.length) return emptyStore(session.renderer.resolve(session.locale));
  const xml = parts.items[0].getXml();
  await context.sync();
  return storeFromXml(xml.value);
}

/** Citations and bibliographies in document order, with their texts. */
async function snapshot(context) {
  const all = citationControls(context);
  all.load("items/tag,items/title,items/id,items/placeholderText");
  await context.sync();
  const ccs = all.items.filter(isOurs);
  const cur = ccs.map((cc) => cc.getReviewedText("Current"));
  const orig = ccs.map((cc) => (isCite(cc) ? cc.getReviewedText("Original") : null));
  await context.sync();
  // Word wraps a control's reviewed text in "<…>". An emptied control shows its placeholder.
  const text = (cc, r) => {
    const v = r.value.startsWith("<") && r.value.endsWith(">") ? r.value.slice(1, -1) : r.value;
    return v === cc.placeholderText ? "" : v;
  };
  const controls = ccs.map((cc, i) => (isCite(cc)
    ? { kind: "cite", tag: cc.tag, title: cc.title, current: text(cc, cur[i]), original: text(cc, orig[i]) }
    : { kind: "bib", current: text(cc, cur[i]) }));
  return { ccs, controls };
}

function setFont(range, run) {
  range.font.italic = !!run.italic;
  range.font.bold = !!run.bold;
  range.font.superscript = !!run.superscript;
  range.font.subscript = !!run.subscript;
}

const plain = (runs) => runs.length === 1 && !runs[0].italic && !runs[0].bold && !runs[0].superscript && !runs[0].subscript;

/** Replaces a citation's text. Plain text keeps the formatting around it. */
function writeCitation(cc, runs) {
  if (plain(runs)) { cc.insertText(runs[0].text, "Replace"); return; }
  runs.forEach((run, i) => setFont(cc.insertText(run.text, i ? "End" : "Replace"), run));
}

function writeBibliography(cc, entries) {
  entries.forEach((runs, i) => {
    let p;
    let rest = runs;
    if (i === 0) {
      // Replacing the whole content leaves one paragraph with the bibliography's style.
      setFont(cc.insertText(runs[0].text, "Replace"), runs[0]);
      p = cc.paragraphs.getFirst();
      rest = runs.slice(1);
    } else {
      p = cc.insertParagraph("", "End");
    }
    for (const run of rest) setFont(p.insertText(run.text, "End"), run);
    p.leftIndent = HANG;
    p.firstLineIndent = -HANG;
  });
}

async function withTracking(context, mode, fn) {
  const doc = context.document;
  doc.load("changeTrackingMode");
  await context.sync();
  const was = doc.changeTrackingMode;
  if (was === mode) return fn();
  doc.changeTrackingMode = mode;
  await context.sync();
  try {
    return await fn();
  } finally {
    doc.changeTrackingMode = was;
    await context.sync();
  }
}

/**
 * Brings numbers and bibliographies up to date. Citations in `fresh` (Word
 * content control ids) were just added by the user and follow the document's
 * track-changes mode. Everything else is derived and is written untracked, so
 * renumbering does not bury the real changes in review.
 */
export async function refresh(context, session, { fresh = [], store } = {}) {
  store = store ?? await openStore(context, session);
  const { ccs, controls } = await snapshot(context);
  const { locale } = store.settings;
  const p = plan(controls, store, session.renderer, { locale, bibEmpty: documentText(locale).bibEmpty });
  if (p.store) await saveStore(context, p.store);
  const own = p.citations.filter((c) => fresh.includes(ccs[c.index].id));
  for (const c of own) writeCitation(ccs[c.index], c.runs);
  await context.sync();
  const rest = p.citations.filter((c) => !own.includes(c));
  if (rest.length || p.bibs.length || p.removes.length) {
    await withTracking(context, "Off", async () => {
      for (const i of p.removes) ccs[i].delete(false);
      for (const c of rest) writeCitation(ccs[c.index], c.runs);
      for (const b of p.bibs) writeBibliography(ccs[b.index], b.entries);
      await context.sync();
    });
  }
  return { ...p, store: p.store ?? store };
}

function tagCitation(cc, id, items, text) {
  cc.tag = CITE_PREFIX + id;
  cc.title = citeTitle(items);
  cc.placeholderText = text.citePlaceholder;
  cc.appearance = "BoundingBox";
}

/**
 * Cites `items` at the cursor. Inside an existing citation the items join it,
 * so a click on [3] and a pick gives [3,4]. With track changes on they become a
 * citation of their own right after it, which review can reject on its own.
 */
export async function insertCitation(context, session, items) {
  const store = await openStore(context, session);
  const sel = context.document.getSelection();
  const host = sel.parentContentControlOrNullObject;
  host.load("tag,title,id");
  context.document.load("changeTrackingMode");
  await context.sync();
  let cc;
  let all = items;
  const inCite = !host.isNullObject && isCite(host);
  if (inCite && context.document.changeTrackingMode === "Off") {
    const before = store.citations[host.tag.slice(CITE_PREFIX.length)] ?? [];
    all = [...before, ...items.filter((x) => !before.some((b) => b.id === x.id))];
    cc = host;
  } else if (inCite) {
    cc = host.getRange("After").insertText(CITE_NEW, "Replace").insertContentControl();
  } else if (!host.isNullObject && host.tag === BIB_TAG) {
    throw new UserError("error.inBibliography");
  } else {
    cc = sel.insertText(CITE_NEW, "End").insertContentControl();
  }
  // A new id each time, so copies of a citation elsewhere keep their items.
  const id = newId();
  tagCitation(cc, id, all, documentText(store.settings.locale));
  cc.load("id");
  await context.sync();
  const next = { ...store, citations: { ...store.citations, [id]: all } };
  await saveStore(context, next);
  cc.getRange("After").select("End");
  await context.sync();
  return refresh(context, session, { fresh: [cc.id], store: next });
}

/**
 * Adds a bibliography for the citations since the previous one. An empty or
 * fully selected paragraph, such as a placeholder in a template, becomes the
 * bibliography; otherwise it goes below the cursor's paragraph. Like its
 * entries, the bibliography is written untracked.
 */
export async function insertBibliography(context, session) {
  const store = await openStore(context, session);
  await withTracking(context, "Off", () => addBibliography(context, documentText(store.settings.locale)));
  return refresh(context, session, { store });
}

async function addBibliography(context, text) {
  const sel = context.document.getSelection();
  const host = sel.parentContentControlOrNullObject;
  host.load("tag");
  sel.load("text");
  const para = sel.paragraphs.getFirst();
  para.load("text,styleBuiltIn");
  await context.sync();
  if (!host.isNullObject && isOurs(host)) throw new UserError("error.inControl");
  const whole = !para.text.trim() || sel.text.trim() === para.text.trim();
  const target = whole ? para : para.insertParagraph("", "After");
  if (!whole && /^Heading/.test(para.styleBuiltIn)) target.styleBuiltIn = Word.Style.normal;
  target.load("isListItem");
  await context.sync();
  // The entries bring their own numbers.
  if (target.isListItem) target.detachFromList();
  target.insertText(text.bibEmpty, "Replace");
  const cc = target.insertContentControl();
  cc.tag = BIB_TAG;
  cc.title = text.bibTitle;
  cc.appearance = "BoundingBox";
  await context.sync();
}

// Word's literal search takes at most this many characters, and "^" starts a
// special code such as "^p" in it, so a caret is searched as "^^".
const SEARCH_MAX = 255;
const searchText = (s) => s.replace(/\^/g, "^^");
// Where a search text may be cut: not inside a surrogate pair, and not before
// a combining mark or a joiner, which belong to the character before them.
const ATTACHES = /[\p{M}\u200c\u200d]|[\udc00-\udfff]/u;

/** The longest piece of `text` from its start (or its end) that one search can take. */
function piece(text, fromEnd) {
  let n = 0;
  let len = 0;
  for (const c of fromEnd ? [...text].reverse() : text) {
    // Word counts UTF-16 units, and the caret is searched doubled.
    len += c === "^" ? 2 : c.length;
    if (len > SEARCH_MAX) break;
    n += c.length;
  }
  if (fromEnd) { while (n > 0 && ATTACHES.test(text[text.length - n])) n--; return text.slice(text.length - n); }
  while (n > 0 && ATTACHES.test(text[n])) n--;
  return text.slice(0, n);
}

// How a range compares with a content control's range: inside it, or clear of it.
// Any other relation means that the two overlap.
const INSIDE = new Set(["Inside", "InsideStart", "InsideEnd", "Equal"]);
const APART = new Set(["Before", "After", "AdjacentBefore", "AdjacentAfter", "Unrelated"]);

/** Where `needle` starts in `text`, without overlaps and in order, as Word's search reports it. */
function occurrences(text, needle) {
  const at = new Map();
  for (let i = text.indexOf(needle); i !== -1; i = text.indexOf(needle, i + needle.length)) at.set(i, at.size);
  return at;
}

/**
 * Finds the typed `@key` markers: [{range, items, text}], and in `lost` the
 * texts of markers that the search could not pin down.
 *
 * Word's wildcard search cannot do this. There "@" is an operator, and a
 * pattern cannot say "at the start of a word". So a plain search first tells
 * whether the document has an "@" at all. If it has, model.findMarkers reads
 * each paragraph's text, and a plain search inside the paragraph gives each
 * group's range: one search for the group's text, or for a group longer than
 * a search can be, one for its start and one for its end and the range
 * between them. Each distinct text is searched once per paragraph, and the
 * group is the hit at its place among the text's occurrences. A group whose
 * range does not read exactly as the paragraph's text does is lost, and so is
 * one that reaches into or over a citation or a bibliography, so nothing is
 * replaced on a guess. All this takes four syncs.
 */
export async function locateMarkers(context) {
  const body = context.document.body;
  const probe = body.search("@", { matchCase: true });
  probe.load("items");
  await context.sync();
  if (!probe.items.length) return { markers: [], lost: [] };

  const paragraphs = body.paragraphs;
  paragraphs.load("items/text");
  const controls = citationControls(context);
  controls.load("items/tag,items/text");
  await context.sync();
  const lost = [];
  const groups = [];
  for (const p of paragraphs.items) {
    if (!p.text.includes("@")) continue;
    const searches = new Map();
    const find = (text, at) => {
      let s = searches.get(text);
      if (!s) {
        s = { at: occurrences(p.text, text), hits: p.search(searchText(text), { matchCase: true }) };
        s.hits.load("items");
        searches.set(text, s);
      }
      return { s, nth: s.at.get(at) };
    };
    for (const g of findMarkers(p.text)) {
      if (searchText(g.text).length <= SEARCH_MAX) { groups.push({ g, ends: [find(g.text, g.start)] }); continue; }
      const head = piece(g.text, false);
      const tail = piece(g.text, true);
      if (!head || !tail) { lost.push(g.text); continue; }
      groups.push({ g, ends: [find(head, g.start), find(tail, g.end - tail.length)] });
    }
  }
  // An "@" inside a citation or a bibliography, at any depth, is part of its text.
  const ours = controls.items.filter((cc) => isOurs(cc) && cc.text.includes("@")).map((cc) => cc.getRange("Whole"));
  await context.sync();

  const found = [];
  for (const { g, ends } of groups) {
    // The paragraph's text and the search must agree on where the marker is.
    if (ends.some(({ s, nth }) => nth === undefined || s.hits.items.length !== s.at.size)) { lost.push(g.text); continue; }
    const [head, tail] = ends.map(({ s, nth }) => s.hits.items[nth]);
    const range = tail ? head.expandTo(tail) : head;
    range.load("text");
    found.push({ range, g, inside: ours.map((r) => range.compareLocationWith(r)) });
  }
  await context.sync();
  const markers = [];
  for (const { range, g, inside } of found) {
    if (range.text !== g.text) lost.push(g.text);
    else if (inside.every((x) => APART.has(x.value))) markers.push({ range, items: g.items, text: g.text });
    else if (!inside.some((x) => INSIDE.has(x.value))) lost.push(g.text);
  }
  return { markers, lost };
}
/**
 * Turns typed `@key` markers into citations and brings the document up to
 * date. A marker with an unknown key stays as text and comes back in
 * `unknownMarkers`.
 */
export async function convertMarkers(context, session) {
  const store = await openStore(context, session);
  const { markers, lost } = await locateMarkers(context);
  const text = documentText(store.settings.locale);
  const citations = { ...store.citations };
  const made = [];
  const unknown = [];
  for (const m of markers) {
    const missing = m.items.filter((x) => !store.sources[x.id]).map((x) => x.id);
    if (missing.length) { unknown.push({ text: m.text, keys: missing }); continue; }
    const id = newId();
    const cc = m.range.insertText(CITE_NEW, "Replace").insertContentControl();
    tagCitation(cc, id, m.items, text);
    cc.load("id");
    citations[id] = m.items;
    made.push(cc);
  }
  await context.sync();
  const next = { ...store, citations };
  if (made.length) await saveStore(context, next);
  const res = await refresh(context, session, { fresh: made.map((cc) => cc.id), store: next });
  return { ...res, converted: made.length, unknownMarkers: unknown, lostMarkers: lost };
}

/** Adds sources to the document. `merge(sources)` returns {library, added, updated}. */
export async function addSources(context, session, merge) {
  const store = await openStore(context, session);
  const r = merge(store.sources);
  const next = { ...store, sources: r.library };
  await saveStore(context, next);
  return { ...r, store: next };
}

/** Sets the language of the document's citations and bibliographies. */
export async function setLocale(context, session, locale) {
  const store = await openStore(context, session);
  const next = { ...store, settings: { ...store.settings, locale: session.renderer.resolve(locale) } };
  await saveStore(context, next);
  return refresh(context, session, { store: next });
}

/**
 * A cheap fingerprint of the content controls' order, to notice deletions and
 * pastes. It takes every tag as it is, so it needs no knowledge of formats.
 */
export async function signature(context) {
  const all = citationControls(context);
  all.load("items/tag");
  await context.sync();
  return all.items.map((cc) => cc.tag).join("|");
}
