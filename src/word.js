// Reads and writes the document through Office.js. All decisions are in model.js.
//
// Every operation takes a session: {renderer, locale}. `locale` is the
// language a new document gets for its citations and bibliographies.
/* global Word */

import {
  NS, CITE_PREFIX, BIB_TAG, CITE_NEW,
  emptyStore, storeToXml, storeFromXml, newId, citeTitle, findMarkers, mayTouch, plan, paragraphsOf,
} from "./model.js";
import { documentText, UserError } from "./messages.js";
import { controlTexts, paragraphStyle, OoxmlError } from "./ooxml.js";

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

// The store's custom XML part, as openStore found it in a request context.
const storePart = new WeakMap();

/**
 * Queues the store's save. It is written with the next sync; queue it with
 * other writes to save a round trip. `save` syncs at once.
 */
async function queueStore(context, store) {
  let part = storePart.get(context);
  if (part === undefined) {
    const parts = context.document.customXmlParts.getByNamespace(NS);
    parts.load("items");
    await context.sync();
    part = parts.items[0] ?? null;
  }
  if (part) part.setXml(storeToXml(store));
  else storePart.set(context, context.document.customXmlParts.add(storeToXml(store)));
}

async function saveStore(context, store) {
  await queueStore(context, store);
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
  storePart.set(context, parts.items[0]);
  const xml = parts.items[0].getXml();
  await context.sync();
  return storeFromXml(xml.value);
}

/**
 * Citations and bibliographies in document order, with their texts
 * `current` and `original`: with tracked changes accepted and rejected.
 *
 * Without tracked changes a control's `text` is both, and that is all there
 * is to read. With tracked changes, `text` follows how Word shows them ("All
 * Markup" keeps deleted text in it, "Simple Markup" leaves it out), and
 * getReviewedText costs about 90 ms per control in a large document and
 * gives a moved text both ways. So then the texts come from the body's OOXML,
 * which says what is inserted, deleted and moved, in one read.
 *
 * Throws an OoxmlError, before anything is written, when the OOXML is not
 * well-formed or its controls are not the ones Word lists, in number and in
 * the order of their tags: then the texts cannot be trusted, and a citation
 * read as empty would be removed.
 */
export async function readControls(context) {
  const all = citationControls(context);
  all.load("items/tag,items/title,items/id,items/text");
  await context.sync();
  let tracked = true;
  try {
    const changes = context.document.body.getTrackedChanges();
    changes.load("items");
    await context.sync();
    tracked = changes.items.length > 0;
  } catch {
    // Word cannot list a document's tracked changes when one of them is a move.
  }
  const ccs = all.items.filter(isOurs);
  // An emptied control shows its placeholder: like the reviewed texts below,
  // that reads as no text.
  const shown = (cc) => (cc.placeholderText && cc.text === cc.placeholderText ? "" : cc.text);
  let texts = ccs.map((cc) => ({ current: shown(cc), original: shown(cc) }));
  if (tracked) {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    const ours = controlTexts(ooxml.value).filter((c) => isOurs(c));
    if (ours.length !== ccs.length || ours.some((c, i) => c.tag !== ccs[i].tag)) {
      throw new OoxmlError("error.ooxmlMismatch", { listed: ccs.length, read: ours.length });
    }
    texts = ours;
  }
  const controls = ccs.map((cc, i) => (isCite(cc)
    ? { kind: "cite", tag: cc.tag, title: cc.title, current: texts[i].current, original: texts[i].original }
    : { kind: "bib", current: texts[i].current }));
  return { ccs, controls, all: all.items };
}

function setFont(range, run) {
  range.font.italic = !!run.italic;
  range.font.bold = !!run.bold;
  range.font.superscript = !!run.superscript;
  range.font.subscript = !!run.subscript;
}

const plain = (runs) => runs.every((r) => !r.italic && !r.bold && !r.superscript && !r.subscript);
const textOf = (runs) => runs.map((r) => r.text).join("");

/**
 * Replaces the text of a citation or a bibliography entry. Plain text is one
 * write; only formatted runs cost a write each. Every write into a citation
 * costs Word about as much as the next (15–20 ms in a document of 4,000
 * citations), so the count of writes is the cost.
 */
function writeRuns(target, runs) {
  if (plain(runs)) { target.insertText(textOf(runs), "Replace"); return; }
  runs.forEach((run, i) => setFont(target.insertText(run.text, i ? "End" : "Replace"), run));
}

const escapeXml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const runXml = (run) => {
  const pr = [run.bold && "<w:b/>", run.italic && "<w:i/>", run.smallCaps && "<w:smallCaps/>", run.superscript && '<w:vertAlign w:val="superscript"/>', run.subscript && '<w:vertAlign w:val="subscript"/>'].filter(Boolean).join("");
  return run.text.split("\t").map((t, i) => (i ? `<w:r>${pr && `<w:rPr>${pr}</w:rPr>`}<w:tab/></w:r>` : "") + (t ? `<w:r>${pr && `<w:rPr>${pr}</w:rPr>`}<w:t xml:space="preserve">${escapeXml(t)}</w:t></w:r>` : "")).join("");
};

/**
 * A bibliography's entries as one OOXML package, each a paragraph with a
 * hanging indent and the paragraph style `style` of the bibliography's first
 * paragraph (its style, not its list, indent or run properties), so its style
 * stays. Word takes 300 entries in one piece in
 * a quarter of a second, against about 6 ms for each single entry.
 */
function entriesOoxml(entries, style) {
  const props = `<w:pPr>${style ? `<w:pStyle w:val="${escapeXml(style)}"/>` : ""}<w:ind w:left="${Math.round(HANG * 20)}" w:hanging="${Math.round(HANG * 20)}"/></w:pPr>`;
  const body = entries.map((runs) => `<w:p>${props}${runs.map(runXml).join("")}</w:p>`).join("");
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}

/**
 * Reads what writing the bibliographies `bibs` needs, in one sync, and
 * returns a function that queues the writes: entry by entry where the plan
 * has hunks and the bibliography still has the paragraphs the plan was made
 * from, otherwise the whole bibliography in one piece.
 */
async function prepareBibliographies(context, bibs) {
  const all = bibs.map((b) => {
    const cc = context.document.body.contentControls.getByIdOrNullObject(b.id);
    const ps = cc.paragraphs;
    ps.load("items/text");
    return { b, cc, ps, first: cc.paragraphs.getFirst().getOoxml() };
  });
  if (all.length) await context.sync();
  // Read before anything is written: an unreadable paragraph stops the refresh.
  for (const x of all) x.style = x.cc.isNullObject ? null : paragraphStyle(x.first.value);
  return () => {
    for (const { b, cc, ps, style } of all) {
      if (cc.isNullObject) continue;
      if (!b.hunks || ps.items.length !== b.paragraphs) {
        cc.getRange("Content").insertOoxml(entriesOoxml(b.entries, style), "Replace");
        continue;
      }
      for (const h of [...b.hunks].reverse()) {
        const keep = Math.min(h.remove, h.insert.length);
        for (let t = 0; t < keep; t++) writeRuns(ps.items[h.at + t], b.entries[h.insert[t]]);
        // A paragraph added before the first or after the last one goes in
        // through the control, so that it stays inside it.
        const extra = h.insert.slice(keep).map((j) => b.entries[j]);
        const after = h.at + keep - 1;
        const add = (where) => (runs) => {
          const q = where === "After" ? ps.items[after].insertParagraph("", "After") : cc.insertParagraph("", where);
          writeRuns(q, runs);
          q.leftIndent = HANG;
          q.firstLineIndent = -HANG;
        };
        if (after === ps.items.length - 1) extra.forEach(add("End"));
        else if (after < 0) [...extra].reverse().forEach(add("Start"));
        else [...extra].reverse().forEach(add("After"));
        for (let t = h.at + keep; t < h.at + h.remove; t++) ps.items[t].delete();
      }
    }
  };
}

/**
 * Queues `write` with track changes off and the mode restored in the same
 * batch, so nothing the person types meanwhile goes untracked.
 */
async function writeUntracked(context, write) {
  const doc = context.document;
  doc.load("changeTrackingMode");
  await context.sync();
  const was = doc.changeTrackingMode;
  if (was !== "Off") doc.changeTrackingMode = "Off";
  try {
    write();
  } finally {
    if (was !== "Off") doc.changeTrackingMode = was;
    await context.sync();
  }
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

/** Citations per batch while numbers change: about half a second of Word's time in a large document. */
export const CHUNK = 30;

/**
 * Brings numbers and bibliographies up to date. Citations in `fresh` (Word
 * content control ids) were just added by the person and follow the
 * document's track-changes mode. Everything else is derived and is written
 * untracked, so renumbering does not bury the real changes in review.
 *
 * It writes only what differs: removed leftovers, the store, the
 * bibliographies and the first `CHUNK` citations. What is left comes back as
 * `work`, for writeChunk; the pane writes it in batches while it stays
 * usable. A document left half written, by a closed pane or document, is put
 * right by the next refresh, which plans from what the document holds.
 */
export async function refresh(context, session, { fresh = [], store } = {}) {
  store = store ?? await openStore(context, session);
  const { ccs, controls, all } = await readControls(context);
  const { locale } = store.settings;
  const opts = { locale, bibEmpty: documentText(locale).bibEmpty };
  let p = plan(controls, store, session.renderer, opts);
  if (p.store && Object.keys(p.store.citations).length < Object.keys(store.citations).length) {
    // Before stored citations go, keep those of controls in footnotes, headers and text boxes.
    const everywhere = context.document.contentControls;
    everywhere.load("items/tag");
    await context.sync();
    const inBody = new Set(ccs.map((cc) => cc.tag));
    const elsewhere = everywhere.items.filter((cc) => isCite(cc) && !inBody.has(cc.tag)).map((cc) => cc.tag.slice(CITE_PREFIX.length));
    if (elsewhere.length) p = plan(controls, store, session.renderer, { ...opts, elsewhere });
  }
  const bibs = p.bibs.map((b) => ({ id: ccs[b.index].id, entries: b.entries, hunks: b.hunks, paragraphs: paragraphsOf(controls[b.index].current).length }));
  const writeBibs = await prepareBibliographies(context, bibs);
  const own = p.citations.filter((c) => fresh.includes(ccs[c.index].id));
  for (const c of own) writeRuns(ccs[c.index], c.runs);
  if (p.store) await saveStore(context, p.store);
  else if (own.length) await context.sync();
  const rest = p.citations.filter((c) => !own.includes(c)).map((c) => ({ id: ccs[c.index].id, runs: c.runs }));
  if (rest.length || bibs.length || p.removes.length) {
    await writeUntracked(context, () => {
      for (const i of p.removes) ccs[i].delete(false);
      writeBibs();
    });
  }
  const work = { citations: rest, total: rest.length, done: 0 };
  if (rest.length) await writeChunk(context, work);
  // The order of the controls as this refresh leaves it, for the watcher.
  const removed = new Set(p.removes.map((i) => ccs[i]));
  const signature = all.filter((cc) => !removed.has(cc)).map((cc) => cc.tag).join("|");
  return { ...p, store: p.store ?? store, work: work.citations.length ? work : null, signature };
}

/**
 * Writes the next batch of `work.citations` untracked, by content control id,
 * and takes them off the list. A control that is gone meanwhile is skipped.
 */
export async function writeChunk(context, work, size = CHUNK) {
  const batch = work.citations.slice(0, size);
  const ccs = batch.map((c) => {
    const cc = context.document.body.contentControls.getByIdOrNullObject(c.id);
    cc.load("id");
    return cc;
  });
  // The controls are found in the same sync that reads the track-changes mode.
  await writeUntracked(context, () => {
    batch.forEach((c, i) => { if (!ccs[i].isNullObject) writeRuns(ccs[i], c.runs); });
  });
  work.citations = work.citations.slice(batch.length);
  work.done += batch.length;
  return work;
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
 *
 * The new citation gets its number first, from the order of the citations'
 * tags, and `onVisible` is called; then refresh brings the rest up to date.
 */
export async function insertCitation(context, session, items, { onVisible } = {}) {
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
  const { locale } = store.settings;
  tagCitation(cc, id, all, documentText(locale));
  cc.load("id");
  const tags = citationControls(context);
  tags.load("items/tag,items/title,items/id");
  await context.sync();
  const next = { ...store, citations: { ...store.citations, [id]: all } };
  // Its number, from the order of the tags alone: texts cost twice as much to read.
  const ours = tags.items.filter(isOurs);
  const at = ours.findIndex((x) => x.id === cc.id);
  const quick = ours.map((x) => (isCite(x) ? { kind: "cite", tag: x.tag, title: x.title, current: CITE_NEW, original: CITE_NEW } : { kind: "bib", current: "" }));
  const first = plan(quick, next, session.renderer, { locale }).citations.find((c) => c.index === at);
  if (first) writeRuns(cc, first.runs);
  cc.getRange("After").select("End");
  // The number, the cursor and the store go to Word in one batch.
  await saveStore(context, next);
  onVisible?.();
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

// Beyond this many "@" in the body, all paragraphs are read at once instead of one per "@".
const PARAGRAPHS_MAX = 400;

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
 * group is the hit at its place among the text's occurrences. Only the
 * paragraphs with an "@" are read, and the controls only when there is a marker. A group whose
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

  // Only the paragraphs that hold an "@" are read: a document with a few
  // e-mail addresses does not pay for reading all of its text.
  let paragraphs;
  if (probe.items.length > PARAGRAPHS_MAX) {
    paragraphs = body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();
    paragraphs = paragraphs.items;
  } else {
    const each = probe.items.map((r) => { const p = r.paragraphs.getFirst(); p.load("text"); return p; });
    const same = each.map((p, i) => (i ? p.getRange("Whole").compareLocationWith(each[i - 1].getRange("Whole")) : null));
    await context.sync();
    paragraphs = each.filter((p, i) => !i || same[i].value !== "Equal");
  }
  const lost = [];
  const groups = [];
  const marked = paragraphs.filter((p) => p.text.includes("@")).map((p) => ({ p, groups: findMarkers(p.text) })).filter((x) => x.groups.length);
  if (!marked.length) return { markers: [], lost: [] };
  const controls = citationControls(context);
  controls.load("items/tag,items/text");
  for (const { p, groups: found } of marked) {
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
    for (const g of found) {
      if (searchText(g.text).length <= SEARCH_MAX) { groups.push({ g, ends: [find(g.text, g.start)] }); continue; }
      const head = piece(g.text, false);
      const tail = piece(g.text, true);
      if (!head || !tail) { lost.push(g.text); continue; }
      groups.push({ g, ends: [find(head, g.start), find(tail, g.end - tail.length)] });
    }
  }
  await context.sync();
  // A marker must not touch a citation or a bibliography, at any depth. Word
  // compares a group's range only with the controls whose text can meet it.
  const ours = controls.items.filter(isOurs);
  const whole = new Map();
  for (const group of groups) {
    group.near = ours.filter((cc) => mayTouch(group.g.text, cc.text));
    for (const cc of group.near) if (!whole.has(cc)) whole.set(cc, cc.getRange("Whole"));
  }

  const found = [];
  for (const { g, ends, near } of groups) {
    // The paragraph's text and the search must agree on where the marker is.
    if (ends.some(({ s, nth }) => nth === undefined || s.hits.items.length !== s.at.size)) { lost.push(g.text); continue; }
    const [head, tail] = ends.map(({ s, nth }) => s.hits.items[nth]);
    const range = tail ? head.expandTo(tail) : head;
    range.load("text");
    found.push({ range, g, inside: near.map((cc) => range.compareLocationWith(whole.get(cc))) });
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
