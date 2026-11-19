// The end-to-end suite's page (dist-e2e/e2e.html). It runs in the task pane
// next to the real pane, drives the pane as a person does, and reports each
// check to e2e/run.mjs. The scenario comes from e2e-config.json.
/* global Word */

import {
  post, $, sleep, command, word, nextDone, done, settledEvents, onReady, idle, key, pick, refresh,
  isCite, isBib, controls, citeTexts, bibTexts, readStore, trackedChanges, setTracking, cursorAfter,
} from "./pane.js";
import { CITE_PREFIX } from "../src/model.js";

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok: !!ok });
  post({ check: name, ok: !!ok, detail: ok ? undefined : detail });
  return !!ok;
}
/**
 * A check that fails because of a defect the add-in has and the README lists
 * as a limit. It is reported as KNOWN, with the reason, and does not fail the
 * run; once it passes it says so.
 */
function known(name, ok, detail, why) {
  results.push({ name, ok: true });
  post({ check: name, ok: !!ok, known: why, detail: ok ? undefined : detail });
  return !!ok;
}
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/** Runs a pane action and waits until the pane has finished it. */
async function act(trigger, timeout) {
  await idle();
  const d = nextDone(timeout);
  trigger();
  const r = await d;
  await idle();
  return r;
}
const cite = (query, opts) => act(() => pick(query, opts));
const insertBibliography = () => act(() => $("bib").click());

/** Imports a library file through the file input. */
async function importFile(text, name) {
  const dt = new DataTransfer();
  dt.items.add(new File([text], name, { type: "text/plain" }));
  $("file").files = dt.files;
  return act(() => $("file").dispatchEvent(new Event("change")));
}

/** Fetches a DOI or PMID through the lookup form and returns the status it shows. */
async function lookup(ident) {
  const before = $("status").textContent;
  $("ident").value = ident;
  const d = nextDone(60000);
  $("lookup").dispatchEvent(new Event("submit", { cancelable: true }));
  await d.catch(() => null);
  // The status follows the operation.
  for (let i = 0; i < 50 && $("status").textContent === before; i++) await sleep(100);
  await idle();
  return $("status").textContent;
}

/** The i-th citation control, by its place in the document. */
const nthCitation = async (c, i) => {
  const cc = c.document.body.contentControls;
  cc.load("items/tag");
  await c.sync();
  return cc.items.filter(isCite)[i];
};

const bodyText = () => word(async (c) => { c.document.body.load("text"); await c.sync(); return c.document.body.text; });

/** Each citation's stored items as keys with their locator or suffix, in document order. */
async function citedKeys() {
  const { store } = await readStore();
  return (await controls()).filter((x) => x.kind === "cite").map((x) => (store.citations[x.tag.slice(CITE_PREFIX.length)] ?? []).map((i) => i.id + (i.locator ? `@${i.locator}` : "") + (i.suffix ? `+${i.suffix}` : "")).join(" "));
}

/**
 * Sets how Word shows tracked changes in this document's window: "all" for
 * All Markup, "simple" for Simple Markup. Returns the mode it had.
 */
const MARKUP = { none: 0, simple: 1, all: 2 };
async function markup(mode) {
  return command(`$v = $doc.ActiveWindow.View.RevisionsFilter; $was = $v.Markup; $v.Markup = ${MARKUP[mode] ?? mode}; $was`);
}

/** Clicks Refresh twice: the second one must write nothing. */
async function settled(name) {
  await refresh();
  const r = await refresh();
  check(`${name}: a second refresh writes nothing`, r && !r.citations.length && !r.bibs.length && !r.removes.length, r && { c: r.citations, b: r.bibs.length, r: r.removes });
}

// ---- the main scenario: an empty document ---------------------------------------

const BIB = String.raw`
@article{mueller2020, author={M{\"u}ller, Anna and Smith, John}, title={Tree rings and summer temperature in the Alps},
  journal={J Ecol}, year={2020}, volume={108}, pages={301--310}, doi={10.1000/jecol.2001}}
@techreport{survey2023, author={{Alpine Survey Office}}, title={Glacier inventory AS-123}, year={2023}, institution={Alpine Survey Office}}
@book{lee2010, author={Lee, Kim}, title={Statistics for field data}, publisher={Springer}, address={Berlin}, year={2010}}
@online{snow2024, author={{Alpine Survey Office}}, title={Snow depth data}, url={https://example.org/snow}, year={2024}, urldate={2026-11-02}}
`;

async function setup() {
  await word(async (c) => {
    const b = c.document.body;
    b.clear();
    c.document.changeTrackingMode = "Off";
    const lines = [
      "Section 1", "First statement.", "Second statement @survey2023 with a marker.", "Third statement.", "References 1",
      "Section 2", "Fourth statement.", "Fifth statement.", "References 2", "Appendix with a statement.",
    ];
    b.insertParagraph(lines[0], "Start");
    for (const l of lines.slice(1)) b.insertParagraph(l, "End");
    const t = b.insertTable(2, 2, "End", [["Site", "Source"], ["AS-123", "Table cell."]]);
    t.styleBuiltIn = Word.Style.gridTable1Light;
    b.insertParagraph("The end.", "End");
    await c.sync();
  });
}

async function main() {
  const fresh = await readStore();
  check("a new document has no store", fresh.parts === 0, fresh);
  await setup();

  await importFile(BIB, "library.bib");
  check("import", /4 new/.test($("status").textContent), $("status").textContent);
  $("q").value = "muel";
  $("q").dispatchEvent(new Event("input"));
  check("fuzzy search", document.querySelector("#results li .key")?.textContent === "mueller2020", document.querySelector("#results li .key")?.textContent);

  await cursorAfter("First statement");
  await cite("muel smith");
  check("cite by search", same(await citeTexts(), ["[1]"]), await citeTexts());

  let r = await refresh();
  check("marker converted", same(await citeTexts(), ["[1]", "[2]"]) && r.converted === 1, [await citeTexts(), $("status").textContent]);

  r = await refresh();
  check("idle refresh writes nothing", r && !r.citations.length && !r.bibs.length && !r.removes.length, r && { c: r.citations.length, b: r.bibs.length });

  await cursorAfter("References 1");
  await insertBibliography();
  let bibs = await bibTexts();
  check("first bibliography", bibs.length === 1 && bibs[0].length === 2 && /^1\.\s*Müller A, Smith J\. Tree rings/.test(bibs[0][0]) && /^2\.\s*Alpine Survey Office/.test(bibs[0][1]), bibs);

  await cursorAfter("Third statement");
  pick("lee", { ctrlKey: true });
  pick("snow", { ctrlKey: true });
  await act(() => key($("q"), "Enter"));
  check("multi cite", same(await citeTexts(), ["[1]", "[2]", "[3,4]"]), await citeTexts());

  await cursorAfter("Fourth statement");
  await cite("lee");
  await cursorAfter("References 2");
  await insertBibliography();
  bibs = await bibTexts();
  check("second section restarts", same(await citeTexts(), ["[1]", "[2]", "[3,4]", "[1]"]) && bibs.length === 2 && bibs[0].length === 4 && bibs[1].length === 1 && /^1\.\s*Lee K/.test(bibs[1][0]), [await citeTexts(), bibs]);

  r = await refresh();
  check("idle refresh with bibliographies", r && !r.citations.length && !r.bibs.length, r && { c: r.citations, b: r.bibs.length });

  await cursorAfter("Appendix with a statement");
  await cite("survey2023");
  await cursorAfter("Table cell");
  await cite("mueller2020");
  bibs = await bibTexts();
  check("appendix and table join the last bibliography", same(await citeTexts(), ["[1]", "[2]", "[3,4]", "[1]", "[2]", "[3]"]) && bibs[1].length === 3, [await citeTexts(), bibs]);

  await word(async (c) => { (await nthCitation(c, 3)).getRange("Content").select("End"); await c.sync(); });
  await cite("snow");
  check("a pick inside a citation joins it", (await citeTexts())[3] === "[1,2]", await citeTexts());

  await cursorAfter("Second statement");
  await cite("lee");
  bibs = await bibTexts();
  check("renumber", same((await citeTexts()).slice(0, 4), ["[1]", "[2]", "[3]", "[2,4]"]) && /^2\.\s*Lee K/.test(bibs[0][1]), [await citeTexts(), bibs[0]]);

  await word(async (c) => { (await nthCitation(c, 1)).delete(false); await c.sync(); });
  await refresh();
  bibs = await bibTexts();
  check("delete", same((await citeTexts()).slice(0, 3), ["[1]", "[2]", "[3,4]"]) && bibs[0].length === 4 && /^3\.\s*Lee K/.test(bibs[0][2]), [await citeTexts(), bibs[0]]);

  await word(async (c) => { (await nthCitation(c, 1)).getRange("Content").delete(); await c.sync(); });
  await refresh();
  const after = await citeTexts();
  check("an emptied citation goes", after.length === 5 && after[1] === "[2,3]", after);

  await word(async (c) => {
    const ox = (await nthCitation(c, 0)).getRange("Whole").getOoxml();
    await c.sync();
    const r2 = c.document.body.search("The end", { matchCase: true });
    r2.load("items");
    await c.sync();
    r2.items[0].getRange("End").insertOoxml(ox.value, "Replace");
    await c.sync();
  });
  await refresh();
  const pasted = await citeTexts();
  // The copy sits behind the last bibliography, where this source is number 3.
  check("a pasted copy", pasted.length === 6 && pasted[5] === pasted[4], pasted);

  // Track changes on: the new citation is a tracked insertion, the renumbering
  // and the bibliography are not, and the mode stays on.
  await setTracking("TrackAll");
  await cursorAfter("First statement");
  await cite("snow");
  const mode = await word(async (c) => { c.document.load("changeTrackingMode"); await c.sync(); return c.document.changeTrackingMode; });
  const tracked = await trackedChanges();
  const onlyCitation = tracked.length > 0 && tracked.every(([type, text]) => type === "Added" && /^\[[\d,–-]+\]$/.test(text.trim()));
  check("with track changes on only the new citation is tracked", mode === "TrackAll" && onlyCitation, { mode, tracked });
  const nums = await citeTexts();
  check("tracked renumber", same(nums.slice(0, 3), ["[1]", "[2]", "[1,3]"]) && nums.length === 7, nums);

  await word(async (c) => { (await nthCitation(c, 1)).getRange("Content").select("End"); await c.sync(); });
  await cite("survey2023");
  const joined = await citeTexts();
  check("with track changes on a pick inside a citation is a citation of its own", joined.length === 8 && joined[1] === "[2]" && joined[2] === "[3]", joined);
  check("that citation is only added", (await trackedChanges()).every(([t]) => t === "Added"), await trackedChanges());

  // A tracked deletion of an older citation waits for review, without a
  // number, however Word shows tracked changes.
  const shown = await markup("all");
  await word(async (c) => { (await nthCitation(c, 1)).getRange("Content").delete(); await c.sync(); });
  await refresh();
  const left = await citeTexts();
  check("a tracked deletion waits for review", /1 deleted, waiting for review/.test($("status").textContent) && left.length === 8 && left[0] === "[1]" && left[2] === "[2]" && left[3] === "[1,3]", { status: $("status").textContent, left });
  await markup("simple");
  const again = await refresh();
  check("with Simple Markup shown, the same deletion still waits and nothing is written", again.pending === 1 && !again.citations.length && !again.bibs.length && !again.removes.length, { pending: again.pending, c: again.citations.length, b: again.bibs.length });
  await markup(shown);
  await setTracking("Off");

  check("doi lookup", /^Added: \S+/.test(await lookup("10.1038/171737a0")), $("status").textContent);
  check("pmid lookup", /^Added: altschul1990/.test(await lookup("PMID 2231712")), $("status").textContent);

  // A group with a supplement becomes one citation; an unknown key and an e-mail address stay text.
  const before = (await citeTexts()).length;
  await word(async (c) => { c.document.body.insertParagraph("Typed: @lee2010 @mueller2020[p. 12] and @nokey2000. Write to anna@example.org.", "End"); await c.sync(); });
  r = await refresh();
  const typed = await citeTexts();
  const group = Object.values(r.store.citations).find((items) => items.some((x) => x.locator));
  check("a marker group with a page", typed.length === before + 1 && /^\[\d+[,–-]\d+\]$/.test(typed.at(-1)) && r.converted === 1
    && same(group, [{ id: "lee2010" }, { id: "mueller2020", locator: "12", label: "page" }]), { typed, group, converted: r.converted });
  const text = await bodyText();
  check("an unknown key and an e-mail address stay text", /Unknown: nokey2000/.test($("status").textContent) && text.includes("@nokey2000") && text.includes("anna@example.org") && !text.includes("@lee2010"), { status: $("status").textContent, text: text.slice(-200) });
  await settled("main");
}

// The main scenario's document, saved and opened again.
async function mainReopen() {
  const f = done.first;
  check("the first refresh writes nothing", f && !f.citations.length && !f.bibs.length && !f.removes.length, f && { c: f.citations, b: f.bibs.length, r: f.removes });
  check("sources and citations are there", f && Object.keys(f.store.sources).length === 6 && f.cited === 8 && f.pending === 1, f && { sources: Object.keys(f.store.sources), cited: f.cited, pending: f.pending });
  await settled("reopened");
}

// ---- a template ---------------------------------------------------------------------

const PLACEHOLDER = "[Your text]";
const HEADINGS = [
  ["1 Background", 1], ["1.1 Earlier work", 2], ["References for section 1", 2],
  ["2 Methods", 1], ["References for section 2", 2],
];

async function setupTemplate() {
  await word(async (c) => {
    const b = c.document.body;
    b.clear();
    c.document.changeTrackingMode = "Off";
    await c.sync();
    HEADINGS.forEach(([text, level], i) => {
      const h = i ? b.insertParagraph(text, "End") : b.insertParagraph(text, "Start");
      h.styleBuiltIn = level === 1 ? Word.Style.heading1 : Word.Style.heading2;
      const p = b.insertParagraph(PLACEHOLDER, "End");
      p.styleBuiltIn = text === "References for section 1" ? Word.Style.quote : Word.Style.normal;
      // The last placeholder is a list item, as templates often have them.
      if (i === HEADINGS.length - 1) p.startNewList();
    });
    await c.sync();
  });
}

async function selectPlaceholderAfter(heading) {
  await word(async (c) => {
    const hs = c.document.body.search(heading, { matchCase: true });
    const ps = c.document.body.search(PLACEHOLDER, { matchCase: true });
    hs.load("items");
    ps.load("items");
    await c.sync();
    const h = hs.items[hs.items.length - 1];
    const where = ps.items.map((p) => p.compareLocationWith(h));
    await c.sync();
    ps.items[where.findIndex((w) => w.value === "After" || w.value === "AdjacentAfter")].select();
    await c.sync();
  });
}

async function fill(heading, text) {
  await selectPlaceholderAfter(heading);
  await word(async (c) => { c.document.getSelection().insertText(text, "Replace").select("End"); await c.sync(); });
}

async function template() {
  await setupTemplate();
  await importFile(BIB, "library.bib");
  check("import", /4 new/.test($("status").textContent), $("status").textContent);

  await fill("1 Background", "Tree rings record summer temperature.");
  await cite("mueller");
  await fill("1.1 Earlier work", "A survey mapped the glaciers.");
  await cite("survey2023");
  await selectPlaceholderAfter("References for section 1");
  await insertBibliography();

  await fill("2 Methods", "The method is standard.");
  pick("lee", { ctrlKey: true });
  await cite("mueller");
  await setTracking("TrackAll");
  await selectPlaceholderAfter("References for section 2");
  await insertBibliography();
  const tracked = await trackedChanges();
  await setTracking("Off");
  check("a bibliography inserted with track changes on is not tracked", tracked.length === 0, tracked);

  const cites = await citeTexts();
  const bibs = await bibTexts();
  check("one bibliography per section", same(cites, ["[1]", "[2]", "[1,2]"]) && bibs.length === 2 && bibs[0].length === 2 && bibs[1].length === 2 && /^1\.\s*Lee K/.test(bibs[1][0]), { cites, bibs });
  const styles = await word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag");
    await c.sync();
    const ps = all.items.filter(isBib).map((x) => { const p = x.paragraphs; p.load("items/styleBuiltIn,items/isListItem"); return p; });
    await c.sync();
    return ps.flatMap((p) => p.items.map((q) => [q.styleBuiltIn, q.isListItem]));
  });
  check("a bibliography in a heading or list placeholder is a plain paragraph", styles.length === 4 && styles.every(([b, li]) => !/^Heading/.test(b) && !li), styles);
  check("a bibliography keeps its placeholder's paragraph style", styles.slice(0, 2).every(([b]) => b === "Quote"), styles.slice(0, 2));
}

async function templateReopen() {
  const f = done.first;
  check("the first refresh writes nothing", f && !f.citations.length && !f.bibs.length && !f.removes.length, f && { c: f.citations, b: f.bibs.length, r: f.removes });
  check("sources and citations are there", f && f.cited === 3 && Object.keys(f.store.sources).length === 4, f && { cited: f.cited, unknown: f.unknown });
  await cursorAfter("The method is standard");
  await cite("snow");
  const bibs = await bibTexts();
  check("citing goes on", same(await citeTexts(), ["[1]", "[2]", "[1]", "[2,3]"]) && bibs[1].length === 3 && /^1\.\s*Alpine Survey Office\. Snow depth data/.test(bibs[1][0]), { cites: await citeTexts(), bibs });
  await settled("reopened");
}

// ---- typed markers in hard places -----------------------------------------------------

const work = (id, family, title) => ({ id, type: "article-journal", author: [{ family, given: "A" }], title, "container-title": "Test J", issued: { "date-parts": [[2020]] } });
// 130 letters from CJK Extension B: 260 UTF-16 units, more than one search takes.
const WIDE = Array.from({ length: 130 }, (_, i) => String.fromCodePoint(0x20000 + i)).join("");
const LONG = Array.from({ length: 25 }, (_, i) => `source${String(i + 1).padStart(2, "0")}`);
const MARKER_SOURCES = [
  work("a", "Alpha", "First work"), work("b", "Beta", "Second work"), work("c", "Gamma", "Third work"),
  work("longkey2020", "Delta", "Fourth work"), work("café", "Epsilon", "Fifth work"), work(WIDE, "Zeta", "Sixth work"),
  ...LONG.map((k, i) => work(k, "Long", `Work number ${i + 1} of the long group`)),
];

/** Clears the body and writes `paragraphs`: strings, or functions (c, body) that add content. */
async function body(...paragraphs) {
  await word(async (c) => {
    const b = c.document.body;
    b.clear();
    c.document.changeTrackingMode = "Off";
    await c.sync();
    let first = true;
    for (const p of paragraphs) {
      if (typeof p === "string") b.insertParagraph(p, first ? "Start" : "End");
      else await p(c, b, first);
      first = false;
      await c.sync();
    }
  });
}

/** The OOXML of the body, for counting fields and comments. */
const bodyOoxml = () => word(async (c) => { const x = c.document.body.getOoxml(); await c.sync(); return x.value; });
const NUMBER = /\[\d+(?:[,–]\d+)*(?:, [^\]]*)?\]/g;
/** The body's text with each citation number as "[#]". */
const shape = async () => (await bodyText()).replace(NUMBER, "[#]").replace(/\r+$/, "");

/**
 * One marker case: writes the body, refreshes and compares the body's text,
 * the citations' stored keys and the markers reported as not converted.
 */
async function markerCase(name, paragraphs, want, extra, why) {
  await body(...paragraphs);
  const r = await refresh();
  const got = { text: await shape(), cites: await citedKeys(), lost: r.lostMarkers ?? [], unknown: (r.unknownMarkers ?? []).map((u) => u.text) };
  const more = extra ? await extra(r) : true;
  const ok = (want.text == null || got.text === want.text) && same(got.cites, want.cites) && same(got.lost, want.lost ?? []) && same(got.unknown, want.unknown ?? []) && more === true;
  (why ? (...a) => known(...a, why) : check)(name, ok, { got, want, more });
}

/** A paragraph with a DATE field before and after the markers. */
const withFields = async (c, b) => {
  const p = b.insertParagraph("Before ", "Start");
  await c.sync();
  const fld = (instr) => `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>FIELD</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  p.getRange("End").insertOoxml(fld("QUOTE one"), "Replace");
  await c.sync();
  p.getRange("End").insertText(" @a and @b ", "Replace");
  await c.sync();
  p.getRange("End").insertOoxml(fld("QUOTE two"), "Replace");
  await c.sync();
  p.getRange("End").insertText(" after.", "Replace");
};

/** Adds a paragraph with a content control of another add-in around `inner`. */
const foreign = (before, inner, after, { cite: citeInside = false } = {}) => async (c, b, first) => {
  const p = b.insertParagraph(before, first ? "Start" : "End");
  const r = p.getRange("End").insertText(inner, "Replace");
  const cc = r.insertContentControl();
  cc.tag = "other-addin:box";
  cc.title = "Another add-in";
  if (citeInside) cc.getRange("End").insertText(" @c", "Replace");
  p.getRange("End").insertText(after, "Replace");
};

/** How many comments the body holds. */
const commentCount = () => word(async (c) => { const cs = c.document.body.getComments(); cs.load("items"); await c.sync(); return cs.items.length; });

async function markers() {
  await importFile(JSON.stringify(MARKER_SOURCES), "markers.json");
  check("import", new RegExp(`${MARKER_SOURCES.length} new`).test($("status").textContent), $("status").textContent);

  // Paragraph text and literal search.
  await markerCase("a field before and after the markers", [withFields], { cites: ["a", "b"] }, async () => {
    const x = await bodyOoxml();
    const n = (x.match(/w:fldCharType="begin"/g) ?? []).length;
    const t = await bodyText();
    return n === 2 && /Before .*\[\d+\] and \[\d+\] .* after\./.test(t) ? true : { fields: n, text: t };
  });
  await markerCase("a tracked insertion that holds the marker", [async (c, b) => {
    b.insertParagraph("Plain text ", "Start");
    await c.sync();
    c.document.changeTrackingMode = "TrackAll";
    await c.sync();
    b.paragraphs.getFirst().getRange("End").insertText("inserted @a here", "Replace");
    await c.sync();
    c.document.changeTrackingMode = "Off";
  }], { text: "Plain text inserted [#] here", cites: ["a"] });
  await markerCase("a tracked deletion just before the marker", [async (c, b) => {
    const p = b.insertParagraph("Keep gone @b.", "Start");
    await c.sync();
    c.document.changeTrackingMode = "TrackAll";
    await c.sync();
    const g = p.search("gone ", { matchCase: true });
    g.load("items");
    await c.sync();
    g.items[0].delete();
    await c.sync();
    c.document.changeTrackingMode = "Off";
  }], { cites: ["b"] }, async () => {
    const tc = await trackedChanges();
    const t = await bodyText();
    return tc.length === 1 && tc[0][0] === "Deleted" && /\[\d+\]\./.test(t) ? true : { tc, t };
  });
  await markerCase("a no-break space between two markers: one citation", ["See @a @b here."], { text: "See [#] here.", cites: ["a b"] });
  await markerCase("a no-break space before a marker", ["See @a here."], { text: "See [#] here.", cites: ["a"] });
  await markerCase("a tab between two markers: two citations", ["See @a\t@b here."], { text: "See [#]\t[#] here.", cites: ["a", "b"] });
  await markerCase("a caret in a supplement", ["See @a[x^2] here."], { text: "See [#] here.", cites: ["a+x^2"] });
  await markerCase("a key typed with a combining accent", ["See @café here."], { text: "See [#] here.", cites: ["café"] });
  await markerCase("a comment anchored on the marker", [async (c, b) => {
    const p = b.insertParagraph("See @a here.", "Start");
    await c.sync();
    const m = p.search("@a", { matchCase: true });
    m.load("items");
    await c.sync();
    m.items[0].insertComment("A note on the marker.");
  }], { text: "See [#] here.", cites: ["a"] });
  // Word deletes a comment when all of its text goes, and a marker's text is
  // replaced by the citation: the comment may not survive, but the conversion
  // above must still be exact.
  const comments = await commentCount();
  known("a comment anchored on the marker survives", comments === 1, { comments },
    "Word deletes a comment when all of its text goes, and a marker's text is replaced by the citation.");

  // Places.
  await markerCase("markers in table cells, the same marker twice in a row", [async (c, b) => {
    b.insertParagraph("Table:", "Start");
    b.insertTable(2, 2, "End", [["@a", "@a"], ["@b here", "none"]]);
    b.insertParagraph("After.", "End");
  }], { cites: ["a", "a", "b"] }, async () => {
    const t = await bodyText();
    return (t.match(NUMBER) ?? []).length === 3 && !t.includes("@") ? true : t;
  });
  await body("One @a here.", "References.");
  await refresh();
  await cursorAfter("References.");
  await insertBibliography();
  const cited = async (c) => {
    const cc = await nthCitation(c, 0);
    cc.getRange("End").insertText(" @b", "Replace");
    const bib = c.document.body.contentControls;
    bib.load("items/tag");
    await c.sync();
    const inBib = bib.items.find(isBib).search("First work", { matchCase: true });
    inBib.load("items");
    await c.sync();
    inBib.items[0].getRange("After").insertText(" @c", "Replace");
    await c.sync();
  };
  await word(cited);
  let r = await refresh();
  check("a marker typed inside a citation or a bibliography entry stays text", r.converted === 0 && same(await citedKeys(), ["a"]) && (await citeTexts())[0] === "[1]", { converted: r.converted, keys: await citedKeys(), cites: await citeTexts(), bibs: await bibTexts() });
  await word(async (c) => {
    const bib = c.document.body.contentControls;
    bib.load("items/tag");
    await c.sync();
    const inBib = bib.items.find(isBib).search("First work", { matchCase: true });
    inBib.load("items");
    await c.sync();
    const inner = inBib.items[0].getRange("After").insertText(" @b", "Replace").insertContentControl();
    inner.tag = "other-addin:note";
    await c.sync();
  });
  r = await refresh();
  check("a marker in another add-in's control inside a bibliography stays text", r.converted === 0 && same(await citedKeys(), ["a"]), { converted: r.converted, keys: await citedKeys() });

  const editedCitation = (where) => async (c, b) => {
    b.insertParagraph("Text ", "Start");
    await c.sync();
    const p = b.paragraphs.getFirst();
    if (where === "before") p.getRange("End").insertText("@b ", "Replace");
    const cc = p.getRange("End").insertText("[1]", "Replace").insertContentControl();
    cc.tag = CITE_PREFIX + "edited1";
    cc.title = "Cite: c";
    await c.sync();
    cc.insertText("@a", "Replace");
    if (where === "after") p.getRange("End").insertText(" @b", "Replace");
    p.getRange("End").insertText(" end.", "Replace");
  };
  await markerCase("a citation edited to read @a, with @b after it: nothing replaced", [editedCitation("after")], { text: "Text [#] @b end.", cites: ["c"], lost: ["@a @b"] });
  await markerCase("a citation edited to read @a, with @b before it: nothing replaced", [editedCitation("before")], { text: "Text @b [#] end.", cites: ["c"], lost: ["@b @a"] });
  await markerCase("a marker that reaches into a citation: nothing replaced", [async (c, b) => {
    b.insertParagraph("Text @", "Start");
    await c.sync();
    const p = b.paragraphs.getFirst();
    const cc = p.getRange("End").insertText("[1]", "Replace").insertContentControl();
    cc.tag = CITE_PREFIX + "reach1";
    cc.title = "Cite: a";
    await c.sync();
    cc.insertText("c", "Start");
    p.getRange("End").insertText(" end.", "Replace");
  }], { text: "Text @[#] end.", cites: ["a"], lost: ["@c[1]"] });
  await markerCase("markers before, inside and after another add-in's control that holds a citation", [async (c, b) => {
    const p = b.insertParagraph("Before @a, ", "Start");
    await c.sync();
    const outer = p.getRange("End").insertText("inside @b and X", "Replace").insertContentControl();
    outer.tag = "other-addin:box";
    const x = outer.search("X", { matchCase: true });
    x.load("items");
    await c.sync();
    const cc = x.items[0].insertText("[1]", "Replace").insertContentControl();
    cc.tag = CITE_PREFIX + "nested1";
    cc.title = "Cite: longkey2020";
    await c.sync();
    p.getRange("End").insertText(", after @c.", "Replace");
  }], { text: "Before [#], inside [#] and [#], after [#].", cites: ["a", "b", "longkey2020", "c"] });
  await markerCase("markers of different lengths in one paragraph", ["Here @longkey2020 and @a then @b[p. 3] and @a @c end."], { text: "Here [#] and [#] then [#] and [#] end.", cites: ["longkey2020", "a", "b@3", "a c"] });

  // Long groups.
  const group = (keys) => keys.map((k, i) => `@${k}[p. ${i + 100}]`).join(" ");
  const g25 = group(LONG);
  const all25 = LONG.map((k, i) => `${k}@${i + 100}`).join(" ");
  await markerCase("a group of 25 markers, longer than a search", [`Long ${g25} end.`], { text: "Long [#] end.", cites: [all25] });
  await markerCase("the same long group twice in a paragraph", [`Long ${g25} and ${g25} end.`], { text: "Long [#] and [#] end.", cites: [all25, all25] });
  const head = group(LONG.slice(0, 20));
  const other = `${head} @a`;
  await markerCase("two long groups that begin alike", [`One ${g25}, two ${other}.`], { text: "One [#], two [#].", cites: [all25, LONG.slice(0, 20).map((k, i) => `${k}@${i + 100}`).join(" ") + " a"] });
  const tail = LONG.slice(5).map((k, i) => `@${k}[p. ${i + 105}]`).join(" ");
  await markerCase("two long groups that end alike", [`One ${g25}, two @a ${tail}.`], { text: "One [#], two [#].", cites: [all25, "a " + LONG.slice(5).map((k, i) => `${k}@${i + 105}`).join(" ")] });
  await markerCase("a key of 130 letters outside the basic plane", [`Wide @${WIDE} end.`], { text: "Wide [#] end.", cites: [WIDE] });

  // Word's Find dialog keeps "Use wildcards" for the next search.
  await command('$f = $doc.Content.Find; $f.ClearFormatting(); $f.MatchWildcards = $true; [void]$f.Execute("x*y"); $doc.Content.Find.MatchWildcards');
  await markerCase("a supplement with ^p, ^t, ^? and * while Use wildcards is on", ["See @a[^p ^t ^? * x] here."], { text: "See [#] here.", cites: ["a+^p ^t ^? * x"] });
  await command('$f = $doc.Content.Find; $f.ClearFormatting(); $f.MatchWildcards = $false; [void]$f.Execute("x"); $f.MatchWildcards');
}

// ---- citations that move between sections -----------------------------------------------

const SECTION_SOURCES = [work("a", "Alpha", "First work"), work("b", "Beta", "Second work"), work("c", "Gamma", "Third work"), work("d", "Delta", "Fourth work")];

/** Copies or moves the first citation whose stored keys are `keys` to after `text`, through Word's own object model. */
const comCopy = (keys, text, { cut = false } = {}) => command(`
$tag = $null
foreach ($cc in $doc.ContentControls) { if ($cc.Title -eq 'Cite: ${keys}') { $src = $cc; break } }
$whole = $doc.Range($src.Range.Start - 1, $src.Range.End + 1)
$r = $doc.Content
$f = $r.Find; $f.ClearFormatting(); $f.MatchWildcards = $false; $f.MatchCase = $true
if (-not $f.Execute('${text}')) { throw 'not found: ${text}' }
$r.Collapse(0)
$r.FormattedText = $whole.FormattedText
${cut ? "$whole.Delete() | Out-Null" : ""}
$doc.ContentControls.Count`);

/** Moves the paragraph that holds `text` below the next one, as Alt+Shift+Down does. */
const comRelocate = (text) => command(`
$r = $doc.Content
$f = $r.Find; $f.ClearFormatting(); $f.MatchWildcards = $false; $f.MatchCase = $true
if (-not $f.Execute('${text}')) { throw 'not found: ${text}' }
$r.Paragraphs(1).Range.Relocate(1)
$doc.Revisions.Count`);

/**
 * The document's revisions as Word's object model lists them: [[type, text]].
 * Office.js fails on a document with a tracked move, so the checks ask Word directly.
 */
const REVISION = { 1: "Added", 2: "Deleted", 14: "MovedFrom", 15: "MovedTo" };
async function revisions() {
  const v = await command('($doc.Revisions | ForEach-Object { "$($_.Type)`t$($_.Range.Text)" }) -join "`n"');
  return String(v ?? "").split("\n").filter(Boolean).map((l) => { const [t, ...x] = l.split("\t"); return [REVISION[t] ?? t, x.join("\t")]; });
}
const acceptAll = () => command("$doc.Revisions.AcceptAll(); $doc.Revisions.Count");

/** Moves the cursor and waits for the selection watcher's refresh, if any. */
async function watcher(where) {
  const d = nextDone(10000).then((x) => x, () => null);
  await cursorAfter(where);
  const r = await d;
  await idle();
  return r;
}

async function sectionsDoc() {
  await body("Section 1", "One @a.", "Two @b.", "References 1", "Section 2", "Three @c.", "Target.", "References 2", "The end.");
  await refresh();
  await cursorAfter("References 1");
  await insertBibliography();
  await cursorAfter("References 2");
  await insertBibliography();
}

const firstWords = async () => (await bibTexts()).map((b) => b.map((l) => /^\d+\.\s*(\w+)/.exec(l)?.[1]));

async function sections() {
  await importFile(JSON.stringify(SECTION_SOURCES), "sections.json");
  await sectionsDoc();
  check("two sections", same(await citeTexts(), ["[1]", "[2]", "[1]"]) && same(await firstWords(), [["Alpha", "Beta"], ["Gamma"]]), { cites: await citeTexts(), bibs: await bibTexts() });

  await comCopy("b", "Target");
  let w = await watcher("The end");
  check("a copy into the next section: the watcher refreshes", w && same(await citeTexts(), ["[1]", "[2]", "[1]", "[2]"]) && same(await firstWords(), [["Alpha", "Beta"], ["Gamma", "Beta"]]), { refreshed: !!w, cites: await citeTexts(), bibs: await firstWords() });

  await comCopy("a", "Three", { cut: true });
  w = await watcher("One");
  check("a cut and paste into the next section", w && same(await citeTexts(), ["[1]", "[1]", "[2]", "[3]"]) && same(await firstWords(), [["Beta"], ["Alpha", "Gamma", "Beta"]]), { refreshed: !!w, cites: await citeTexts(), bibs: await firstWords() });

  await comCopy("c", "Two", { cut: true });
  w = await watcher("The end");
  check("a cut and paste back into the first section", w && same(await citeTexts(), ["[1]", "[2]", "[1]", "[2]"]) && same(await firstWords(), [["Gamma", "Beta"], ["Alpha", "Beta"]]), { refreshed: !!w, cites: await citeTexts(), bibs: await firstWords() });
  await settled("after moving");

  // A move with track changes on, as Alt+Shift+Down makes it: the paragraph
  // "Two" goes below "Three".
  await body("Section 1", "One @a.", "Two @b.", "Three @c.", "References 1", "Section 2", "Four @d.", "References 2");
  await refresh();
  await cursorAfter("References 1");
  await insertBibliography();
  await cursorAfter("References 2");
  await insertBibliography();
  await setTracking("TrackAll");
  const shown = await markup("all");
  const before = await revisions();
  const moved = await comRelocate("Two");
  w = await watcher("Four");
  if (!w) await refresh();
  const types = await revisions();
  // A control's text follows the markup shown; with Simple Markup the moved-away copy reads empty.
  await markup("simple");
  const cites = await citeTexts();
  await markup("all");
  check("a tracked move of a paragraph: numbers follow the new order", same(cites.filter(Boolean), ["[1]", "[2]", "[3]", "[1]"]) && same((await firstWords())[0], ["Alpha", "Gamma", "Beta"]),
    { moved, types, cites, bibs: await bibTexts(), status: $("status").textContent });
  await markup("simple");
  const again = await refresh();
  check("with Simple Markup shown, the move needs no writes", again && !again.citations.length && !again.bibs.length, again && { c: again.citations, b: again.bibs.length });
  await markup(shown);
  await setTracking("Off");
  await refresh();
  // Whatever the relocate added must still be there: the move stays tracked.
  const still = await revisions();
  const added = types.filter(([t, x]) => !before.some(([t2, x2]) => t2 === t && x2 === x));
  const kept = added.length > 0 && added.every(([t, x]) => still.some(([t2, x2]) => t2 === t && x2 === x));
  check("the move stays tracked for review", moved > before.length && kept, { moved, before: before.length, added, still });

  // Tracked changes with a comment, a field, a text box holding a citation
  // control (which Word omits from body.contentControls) and a deleted
  // paragraph mark: the OOXML read must match Word's list under both views.
  await word(async (c) => {
    const r = c.document.body.search("One", { matchCase: true });
    r.load("items");
    await c.sync();
    r.items[0].insertComment("A note in a tracked document.");
    await c.sync();
  });
  await setTracking("TrackAll");
  await command(`
$r = $doc.Content
$f = $r.Find; $f.ClearFormatting(); $f.MatchWildcards = $false; $f.MatchCase = $true
if (-not $f.Execute('Four')) { throw 'not found: Four' }
$r.Collapse(1)
[void]$doc.Fields.Add($r, 33)
$tb = $doc.Shapes.AddTextbox(1, 300, 10, 120, 40)
$tb.TextFrame.TextRange.Text = '[?]'
$cc = $tb.TextFrame.TextRange.ContentControls.Add(1)
$cc.Tag = '${CITE_PREFIX}intextbox'
$cc.Title = 'Cite: d'
$p1 = $doc.Paragraphs(2)
$doc.Range($p1.Range.End - 1, $p1.Range.End).Delete() | Out-Null
$doc.Revisions.Count`);
  await markup("all");
  const ooxmlAll = await refresh();
  await markup("simple");
  const ooxmlSimple = await refresh();
  const ooxmlCites = (await citeTexts()).filter(Boolean);
  await markup(shown);
  check("tracked changes with a comment, field, text box and deleted paragraph mark read cleanly",
    ooxmlAll && !ooxmlAll.error && ooxmlSimple && !ooxmlSimple.error
    && !ooxmlSimple.citations.length && !ooxmlSimple.bibs.length
    && same(ooxmlCites, ["[1]", "[2]", "[3]", "[1]"]),
    { all: ooxmlAll && { err: ooxmlAll.error, c: ooxmlAll.citations?.length }, simple: ooxmlSimple && { err: ooxmlSimple.error, c: ooxmlSimple.citations?.length }, cites: ooxmlCites, status: $("status").textContent });
  await setTracking("Off");
  await command("while ($doc.Shapes.Count -gt 0) { $doc.Shapes(1).Delete() }; $doc.Shapes.Count");
  await word(async (c) => {
    const cs = c.document.body.getComments();
    cs.load("items");
    await c.sync();
    for (const x of cs.items) x.delete();
    await c.sync();
  });

  // Renumbering with track changes on: a new first citation renumbers the
  // rest untracked, and the bibliography too.
  await setTracking("TrackAll");
  await acceptAll();
  await cursorAfter("Section 1");
  await act(() => pick("d"));
  const tracked = await revisions();
  check("renumbering with track changes on tracks only the new citation", tracked.length === 1 && tracked[0][0] === "Added" && same((await citeTexts()).slice(0, 4), ["[1]", "[2]", "[3]", "[4]"]), { tracked, cites: await citeTexts() });
  await setTracking("Off");
  await settled("sections");
}

async function sectionsReopen() {
  const f = done.first;
  check("the first refresh writes nothing", f && !f.citations.length && !f.bibs.length && !f.removes.length, f && { c: f.citations, b: f.bibs.length, r: f.removes });
  const tracked = await revisions();
  check("the tracked insertion is still there", tracked.length === 1 && tracked[0][0] === "Added", tracked);
  await acceptAll();
  const r = await refresh();
  check("after accepting, a refresh writes nothing", r && !r.citations.length && !r.bibs.length, r && { c: r.citations, b: r.bibs.length });
  await settled("reopened");
}

// ---- renumbering in batches -------------------------------------------------------------

/** Sources that no citation uses yet: citing one at the start renumbers the section. */
async function unused() {
  const { store } = await readStore();
  const cited = new Set(Object.values(store.citations).flat().map((x) => x.id));
  return Object.keys(store.sources).filter((k) => !cited.has(k));
}
const cursorAtStart = () => word(async (c) => { c.document.body.paragraphs.getFirst().getNext().getRange("Start").select("Start"); await c.sync(); });
const cursorAtEnd = () => word(async (c) => { c.document.body.paragraphs.getLast().getRange("End").select("End"); await c.sync(); });

/** A refresh after everything has settled writes nothing. */
async function consistent(name) {
  await idle();
  const r = await refresh();
  check(name, r && !r.citations.length && !r.bibs.length && !r.removes.length, r && { c: r.citations.length, b: r.bibs.length, r: r.removes.length, status: $("status").textContent });
}

const PHASE = "siglum.e2e.phase";

async function renumber() {
  const fresh = await unused();
  if (sessionStorage.getItem(PHASE) === "reloaded") {
    sessionStorage.removeItem(PHASE);
    const f = done.first;
    check("after the pane closed half-way, its first refresh writes the rest", f && f.citations.length > 0, f && { c: f.citations.length, b: f.bibs.length });
    await consistent("after the pane closed half-way, the document is consistent again");
    // Now the document closes half-way: the harness saves it while the batches go on.
    await cursorAtStart();
    const d = nextDone();
    pick(fresh[0]);
    await d;
    post({ done: true, ok: results.every((x) => x.ok) });
    await new Promise(() => {});
  }
  await cursorAtStart();
  let d = nextDone();
  pick(fresh[0]);
  let r = await d;
  check("a new first citation is numbered at once, the rest follows in batches with a progress line", (await citeTexts())[0] === "[1]" && r.work && /^Updating citations: \d+ of \d+/.test($("status").textContent), { first: (await citeTexts())[0], work: !!r?.work, status: $("status").textContent });
  // A second action while the batches run.
  await cursorAtEnd();
  pick(fresh[1]);
  await consistent("a second citation during the renumbering leaves a consistent document");

  await cursorAtStart();
  d = nextDone();
  pick(fresh[2]);
  await d;
  // A citation deleted by hand while the batches run.
  await word(async (c) => { (await nthCitation(c, 5)).delete(false); await c.sync(); });
  await consistent("a citation deleted during the renumbering leaves a consistent document");
  const tags = await word(async (c) => {
    const cc = c.document.contentControls;
    cc.load("items/tag");
    await c.sync();
    return new Set(cc.items.filter(isCite).map((x) => x.tag.slice(CITE_PREFIX.length)));
  });
  const { store } = await readStore();
  const stale = Object.keys(store.citations).filter((id) => !tags.has(id));
  check("the stored entries of citations that are gone leave the store", stale.length === 0, stale);

  // An import while the batches run: the note appears above the progress line,
  // the dropped writes plan again, and siglum:settled fires once at the end.
  await cursorAtStart();
  const settledBefore = settledEvents.count;
  d = nextDone();
  pick(fresh[3]);
  r = await d;
  const dImport = nextDone();
  const dt = new DataTransfer();
  dt.items.add(new File([JSON.stringify([work("imported2025", "Zeller", "Imported during renumbering")])], "mid.json", { type: "text/plain" }));
  $("file").files = dt.files;
  $("file").dispatchEvent(new Event("change"));
  await dImport;
  await sleep(0);
  const duringImport = $("status").textContent;
  await idle();
  const settledAfter = settledEvents.count - settledBefore;
  const { store: afterImport } = await readStore();
  check("an import during renumbering keeps the progress line and settles once",
    r?.work && /1 new/.test(duringImport) && /Updating citations: \d+ of \d+/.test(duringImport)
    && settledAfter === 1 && Boolean(afterImport?.sources?.imported2025),
    { work: !!r?.work, duringImport, settledAfter, imported: Boolean(afterImport?.sources?.imported2025), status: $("status").textContent });
  await consistent("an import during the renumbering leaves a consistent document");

  // The pane closes half-way: a reload stands for it.
  await cursorAtStart();
  d = nextDone();
  pick(fresh[4]);
  await d;
  sessionStorage.setItem(PHASE, "reloaded");
  location.reload();
  await new Promise(() => {});
}

async function renumberReopen() {
  const f = done.first;
  check("after the document closed half-way, its first refresh writes the rest", f && f.citations.length > 0, f && { c: f.citations.length });
  await consistent("after the document closed half-way, the document is consistent again");
}

// ---- one body write for a large renumber --------------------------------------------

async function fastpath() {
  const fresh = await unused();
  // A new first citation renumbers all 600 and grows the bibliographies:
  // the bibliographies go first, the citations in one body write.
  await cursorAtStart();
  let d = nextDone();
  pick(fresh[0]);
  const r1 = await d;
  await idle();
  check("a 600-citation renumber numbers the new citation first", (await citeTexts())[0] === "[1]", (await citeTexts()).slice(0, 3));
  check("a 600-citation renumber takes one body write", r1 && r1.fast === true, r1 && { fast: r1.fast, work: !!r1.work });
  await consistent("a 600-citation renumber leaves a consistent document");
  // A text box shape passes the body write byte-identical.
  await command(`$tb = $doc.Shapes.AddTextbox(1, 300, 10, 120, 40)\n$tb.TextFrame.TextRange.Text = 'shape keeps'`);
  d = nextDone();
  pick(fresh[1]);
  const r2 = await d;
  await idle();
  const shape = await command(`($doc.Shapes | ForEach-Object { $_.TextFrame.TextRange.Text }) -join '|'`);
  check("a text box shape passes the renumber unchanged", shape === "shape keeps", shape);
  check("a renumber with a text box takes one body write", r2 && r2.fast === true, r2 && { fast: r2.fast, work: !!r2.work });
  await consistent("a renumber with a text box leaves a consistent document");
  // A person's bookmark stays out of the body write and survives it.
  await command(`[void]$doc.Bookmarks.Add('keepMe', $doc.Paragraphs(30).Range)\n$doc.Bookmarks.Count`);
  const names = await command(`($doc.Bookmarks | ForEach-Object { $_.Name }) -join ','`);
  d = nextDone();
  pick(fresh[2]);
  const r3 = await d;
  await idle();
  const namesAfter = await command(`($doc.Bookmarks | ForEach-Object { $_.Name }) -join ','`);
  check("a person's bookmark survives the renumber", namesAfter === names && /(^|,)keepMe(,|$)/.test(namesAfter), { names, namesAfter });
  check("a renumber with a person's bookmark skips the body write", r3 && r3.fast !== true, r3 && { fast: r3.fast });
  await consistent("a renumber with a person's bookmark leaves a consistent document");
  // A modern comment thread reads cleanly under tracked changes.
  await word(async (c) => {
    const r = c.document.body.search("Section", { matchCase: true });
    r.load("items");
    await c.sync();
    r.items[0].insertComment("A note in a fast document.");
    await c.sync();
  });
  const threaded = await command(`try {\n$c = @($doc.Comments)[@($doc.Comments).Count - 1]\n$c.Replies.Add('A reply in the thread.') | Out-Null\n$c.Replies.Count\n} catch { 'unsupported' }`);
  if (threaded === "unsupported") {
    known("a reply in the comment thread", false, threaded, "Word without modern comment replies");
  } else {
    check("a reply joins the comment thread", Number(threaded) === 1, threaded);
  }
  await setTracking("TrackAll");
  await markup("all");
  const ooxmlAll = await refresh();
  await markup("simple");
  const ooxmlSimple = await refresh();
  const cites = (await citeTexts()).filter(Boolean);
  check("tracked changes with a comment thread read cleanly",
    ooxmlAll && !ooxmlAll.error && ooxmlSimple && !ooxmlSimple.error && cites.length > 0,
    { all: ooxmlAll && { err: ooxmlAll.error }, simple: ooxmlSimple && { err: ooxmlSimple.error }, cites: cites.length });
  await setTracking("Off");
  await word(async (c) => {
    const cs = c.document.body.getComments();
    cs.load("items");
    await c.sync();
    for (const x of cs.items) x.delete();
    await c.sync();
  });
  await consistent("a fast document with a thread read leaves a consistent document");
}

const SCENARIOS = { renumber, "renumber-reopen": renumberReopen, fastpath, main, "main-reopen": mainReopen, template, "template-reopen": templateReopen, markers, sections, "sections-reopen": sectionsReopen };

onReady(async () => {
  const { scenario } = await (await fetch("e2e-config.json", { cache: "no-store" })).json();
  try {
    await SCENARIOS[scenario]();
  } catch (e) {
    check("the scenario ran to its end", false, String(e && (e.stack || e.message || e)));
  }
  return results.every((r) => r.ok);
});
