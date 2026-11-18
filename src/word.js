// Reads and writes the document through Office.js. All decisions are in model.js.
//
// Every operation takes a session: {renderer, locale}. `locale` is the
// language a new document gets for its citations and bibliographies.
/* global Word */

import {
  NS, CITE_PREFIX, BIB_TAG, CITE_NEW,
  emptyStore, storeToXml, storeFromXml, mergeStores, newId, citeTitle, findMarkers, plan,
} from "./model.js";
import { upgradeLegacy } from "./legacy.js";
import { documentText, UserError } from "./messages.js";

// Hanging indent of bibliography entries, in points (1 cm).
const HANG = 28.35;
// Word refuses a search text longer than this.
const SEARCH_MAX = 255;

const isCite = (cc) => cc.tag?.startsWith(CITE_PREFIX);
const isOurs = (cc) => isCite(cc) || cc.tag === BIB_TAG;

/**
 * The content controls that can hold citations and bibliographies. Citations
 * in footnotes and endnotes would be added here.
 */
function citationControls(context) {
  return context.document.body.contentControls;
}

async function loadStore(context) {
  const parts = context.document.customXmlParts.getByNamespace(NS);
  parts.load("items");
  await context.sync();
  if (!parts.items.length) return emptyStore();
  const xml = parts.items[0].getXml();
  await context.sync();
  return storeFromXml(xml.value);
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
 * The document's store. Every operation starts here, so this is the one place
 * that upgrades a document in an old format. A document without a language
 * for its citations gets the session's; the next save keeps it.
 */
async function openStore(context, session) {
  await upgradeLegacy(context, {
    adopt: async (old) => saveStore(context, mergeStores(old, await loadStore(context))),
    untracked: (fn) => withTracking(context, "Off", fn),
  });
  const store = await loadStore(context);
  const locale = session.renderer.resolve(store.settings.locale ?? session.locale);
  return { ...store, settings: { ...store.settings, locale } };
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

/**
 * Finds the typed `@key` markers: [{range, items, text}], and in `lost` the
 * texts of markers that the search could not pin down.
 *
 * Word's wildcard search cannot do this. There "@" is an operator, and a
 * pattern cannot say "at the start of a word". So a plain search first tells
 * whether the document has an "@" at all. If it has, model.findMarkers reads
 * each paragraph's text, and a plain search for a marker's exact text inside
 * its paragraph gives the range.
 */
async function locateMarkers(context) {
  const body = context.document.body;
  const probe = body.search("@", { matchCase: true });
  probe.load("items");
  await context.sync();
  if (!probe.items.length) return { markers: [], lost: [] };

  const paragraphs = body.paragraphs;
  paragraphs.load("items/text");
  await context.sync();
  const lost = [];
  const wanted = [];
  for (const p of paragraphs.items) {
    if (!p.text.includes("@")) continue;
    for (const group of findMarkers(p.text)) {
      // "^" starts a special code in a plain search.
      if (group.text.length > SEARCH_MAX || group.text.includes("^")) { lost.push(group.text); continue; }
      const hits = p.search(group.text, { matchCase: true });
      hits.load("items");
      wanted.push({ group, hits });
    }
  }
  await context.sync();

  const found = [];
  for (const { group, hits } of wanted) {
    // The paragraph's text and the search must agree on where the marker is.
    if (hits.items.length !== group.of) { lost.push(group.text); continue; }
    const range = hits.items[group.nth];
    const host = range.parentContentControlOrNullObject;
    host.load("tag");
    found.push({ range, host, items: group.items, text: group.text });
  }
  await context.sync();
  // An "@" inside a citation or a bibliography is part of its text.
  return { markers: found.filter((m) => m.host.isNullObject || !isOurs(m.host)), lost };
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
