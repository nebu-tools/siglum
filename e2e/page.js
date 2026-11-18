// Runs inside Word's task pane next to the real pane (dist-e2e/e2e.html). Drives
// the pane like a user would and reports each check to the e2e server.
/* global Word */

import { BIB_TAG, CITE_PREFIX, NS } from "../src/model.js";
// A document as the version before the stored format was renamed left it.
import legacyXml from "../test/fixtures/legacy/store.xml";
import legacyDoc from "../test/fixtures/legacy/document.json";

// The checks read the pane in English, whatever language Office has.
try { localStorage.setItem("siglum.language", "en"); } catch { /* The checks on status texts then fail and say so. */ }

const results = [];
let last = null;
let first = null;
document.addEventListener("siglum:done", (e) => { last = e.detail; first = first ?? e.detail; });
const post = (o) => fetch("/log", { method: "POST", body: JSON.stringify(o) }).catch(() => {});
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const $ = (id) => document.getElementById(id);

async function idle() {
  await sleep(150);
  for (let calm = 0; calm < 3; ) {
    await sleep(100);
    calm = document.body.hasAttribute("aria-busy") ? 0 : calm + 1;
  }
}

function check(name, ok, detail) {
  results.push({ name, ok: !!ok, detail });
  post({ check: name, ok: !!ok, detail });
}

const word = (fn) => Word.run(fn);

const isCite = (x) => x.tag.startsWith(CITE_PREFIX);

/** Citations and bibliographies as [kind, text] in document order. */
async function controls() {
  return word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag,items/text");
    await c.sync();
    return all.items.filter((x) => isCite(x) || x.tag === BIB_TAG).map((x) => [x.tag === BIB_TAG ? "bib" : "cite", x.text]);
  });
}
const citeTexts = async () => (await controls()).filter((x) => x[0] === "cite").map((x) => x[1]);
const bibTexts = async () => (await controls()).filter((x) => x[0] === "bib").map((x) => x[1].split(/[\r\n]+/).filter(Boolean));

/** What Word reports for the i-th citation. */
async function probe(i) {
  return word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    const x = cc.items.filter(isCite)[i];
    x.load("text,placeholderText,tag");
    const cur = x.getReviewedText("Current");
    const orig = x.getReviewedText("Original");
    const ox = x.getOoxml();
    await c.sync();
    const body = /<w:body>([\s\S]*)<\/w:body>/.exec(ox.value)?.[1] ?? "";
    return { tag: x.tag, text: x.text, ph: x.placeholderText, cur: cur.value, orig: orig.value, xml: body.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").slice(0, 1500) };
  });
}

async function cursorAfter(text) {
  await word(async (c) => {
    const r = c.document.body.search(text, { matchCase: true });
    r.load("items");
    await c.sync();
    if (!r.items.length) throw new Error("not found: " + text);
    r.items[0].getRange("End").select("End");
    await c.sync();
  });
}

function key(target, k, opts = {}) {
  target.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...opts }));
}

async function pick(query, opts) {
  $("q").value = query;
  $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", opts);
  await idle();
}

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

async function scenario() {
  await setup();

  // Import a .bib through the file input.
  const dt = new DataTransfer();
  dt.items.add(new File([BIB], "library.bib", { type: "text/plain" }));
  $("file").files = dt.files;
  $("file").dispatchEvent(new Event("change"));
  await idle();
  check("import", /4 new/.test($("status").textContent), $("status").textContent);
  $("q").value = "muel";
  $("q").dispatchEvent(new Event("input"));
  check("fuzzy search", document.querySelector("#results li .key")?.textContent === "mueller2020", document.querySelector("#results li .key")?.textContent);

  // Cite by search and Enter.
  await cursorAfter("First statement");
  await pick("muel smith");
  check("cite by search", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]"]), await citeTexts());

  // Typed marker, converted by Refresh.
  $("refresh").click();
  await idle();
  check("marker converted", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]", "[2]"]), [await citeTexts(), $("status").textContent]);

  // A refresh of a current document writes nothing.
  $("refresh").click();
  await idle();
  check("idle refresh writes nothing", last && last.citations.length === 0 && last.bibs.length === 0 && last.removes.length === 0, last && { c: last.citations.length, b: last.bibs.length, p: await probe(0) });

  // Bibliography after "References 1".
  await cursorAfter("References 1");
  $("bib").click();
  await idle();
  let bibs = await bibTexts();
  check("first bibliography", bibs.length === 1 && bibs[0].length === 2 && /^1\.\s*Müller A, Smith J\. Tree rings/.test(bibs[0][0]) && /^2\.\s*Alpine Survey Office/.test(bibs[0][1]), bibs);

  // Mark two with Ctrl+Enter, cite both at once.
  await cursorAfter("Third statement");
  $("q").value = "lee"; $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", { ctrlKey: true });
  $("q").value = "snow"; $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", { ctrlKey: true });
  key($("q"), "Enter");
  await idle();
  check("multi cite", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]", "[2]", "[3,4]"]), await citeTexts());

  // Section 2 starts again at 1, and its bibliography is separate.
  await cursorAfter("Fourth statement");
  await pick("lee");
  await cursorAfter("References 2");
  $("bib").click();
  await idle();
  bibs = await bibTexts();
  check("second section restarts", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]", "[2]", "[3,4]", "[1]"]) && bibs.length === 2 && bibs[0].length === 4 && bibs[1].length === 1 && /^1\.\s*Lee K/.test(bibs[1][0]), [await citeTexts(), bibs]);

  // With two bibliographies in place, another refresh still writes nothing.
  $("refresh").click();
  await idle();
  check("idle refresh with bibliographies", last && last.citations.length === 0 && last.bibs.length === 0, last && { c: last.citations, b: last.bibs.length });

  // Appendix after the last bibliography and a table cell join the last one.
  await cursorAfter("Appendix with a statement");
  await pick("survey2023");
  await cursorAfter("Table cell");
  await pick("mueller2020");
  bibs = await bibTexts();
  check("appendix and table", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]", "[2]", "[3,4]", "[1]", "[2]", "[3]"]) && bibs[1].length === 3, [await citeTexts(), bibs]);

  // Inside an existing citation, a pick joins it.
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    cc.items.filter(isCite)[3].getRange("Content").select("End");
    await c.sync();
  });
  await pick("snow");
  check("join citation", (await citeTexts())[3] === "[1,2]", await citeTexts());

  // A new citation early on renumbers later ones and the bibliography.
  await cursorAfter("Second statement");
  await pick("lee");
  bibs = await bibTexts();
  check("renumber", JSON.stringify((await citeTexts()).slice(0, 4)) === JSON.stringify(["[1]", "[2]", "[3]", "[2,4]"]) && /^2\.\s*Lee K/.test(bibs[0][1]), [await citeTexts(), bibs[0]]);

  // Deleting a citation: the bibliography follows on refresh.
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    cc.items.filter(isCite)[1].delete(false);
    await c.sync();
  });
  $("refresh").click();
  await idle();
  bibs = await bibTexts();
  check("delete", JSON.stringify((await citeTexts()).slice(0, 3)) === JSON.stringify(["[1]", "[2]", "[3,4]"]) && bibs[0].length === 4 && /^3\.\s*Lee K/.test(bibs[0][2]), [await citeTexts(), bibs[0]]);

  // Emptying a citation's text by hand removes the leftover control.
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    cc.items.filter(isCite)[1].getRange("Content").delete();
    await c.sync();
  });
  const emptied = await probe(1);
  $("refresh").click();
  await idle();
  const after = await citeTexts();
  check("emptied", after.length === 5 && after[1] === "[2,3]", { after, emptied });

  // A pasted copy of a citation: same number, both survive.
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    const ox = cc.items.filter(isCite)[0].getRange("Whole").getOoxml();
    await c.sync();
    const r = c.document.body.search("The end", { matchCase: true });
    r.load("items");
    await c.sync();
    r.items[0].getRange("End").insertOoxml(ox.value, "Replace");
    await c.sync();
  });
  $("refresh").click();
  await idle();
  const pasted = await citeTexts();
  // The copy sits behind the last bibliography, where this source is number 3.
  check("paste copy", pasted.length === 6 && pasted[5] === pasted[4], pasted);

  // Track changes on: a new citation is a tracked insertion, the renumbering
  // and the bibliography are not, and the mode stays on.
  await word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); });
  await cursorAfter("First statement");
  await pick("snow");
  const tracked = await word(async (c) => {
    c.document.load("changeTrackingMode");
    const tcs = c.document.body.getTrackedChanges();
    tcs.load("items/type,items/text");
    await c.sync();
    return { mode: c.document.changeTrackingMode, changes: tcs.items.map((t) => [t.type, t.text]) };
  });
  const onlyCitation = tracked.changes.length > 0 && tracked.changes.every(([type, text]) => type === "Added" && /^\[[\d,–-]+\]$/.test(text.trim()));
  check("tracked insert only", tracked.mode === "TrackAll" && onlyCitation, tracked);
  const nums = await citeTexts();
  // The cursor sits before the first citation, so the new one comes first.
  check("tracked renumber", JSON.stringify(nums.slice(0, 3)) === JSON.stringify(["[1]", "[2]", "[1,3]"]) && nums.length === 7, nums);

  // With tracking on, a pick inside a citation becomes its own citation after it.
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    cc.items.filter(isCite)[1].getRange("Content").select("End");
    await c.sync();
  });
  await pick("survey2023");
  const joined = await citeTexts();
  check("tracked join is separate", joined.length === 8 && joined[1] === "[2]" && joined[2] === "[3]", joined);
  const changes = await word(async (c) => {
    const tcs = c.document.body.getTrackedChanges();
    tcs.load("items/type,items/text");
    await c.sync();
    return tcs.items.map((t) => [t.type, t.text]);
  });
  check("tracked join only adds", changes.every(([t]) => t === "Added"), changes);

  // A tracked deletion of an older citation stays for review, without a number.
  // (Deleting one's own tracked insertion just removes it.)
  await word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag");
    await c.sync();
    cc.items.filter(isCite)[1].getRange("Content").delete();
    await c.sync();
  });
  const pend = await probe(1);
  $("refresh").click();
  await idle();
  const left = await citeTexts();
  check("tracked delete pending", /1 deleted, waiting for review/.test($("status").textContent) && left.length === 8 && left[0] === "[1]" && left[2] === "[2]" && left[3] === "[1,3]", { status: $("status").textContent, left, pend });
  await word(async (c) => { c.document.changeTrackingMode = "Off"; await c.sync(); });

  // DOI lookup over the network.
  $("ident").value = "10.1038/171737a0";
  $("lookup").dispatchEvent(new Event("submit", { cancelable: true }));
  await idle();
  check("doi lookup", /^Added: \S+/.test($("status").textContent), $("status").textContent);
  $("ident").value = "PMID 2231712";
  $("lookup").dispatchEvent(new Event("submit", { cancelable: true }));
  await idle();
  check("pmid lookup", /^Added: altschul1990/.test($("status").textContent), $("status").textContent);

  // Typed markers: a group with a supplement becomes one citation, a marker
  // with an unknown key and an e-mail address stay as text.
  const before = (await citeTexts()).length;
  await word(async (c) => {
    c.document.body.insertParagraph("Typed: @lee2010 @mueller2020[p. 12] and @nokey2000. Write to anna@example.org.", "End");
    await c.sync();
  });
  $("refresh").click();
  await idle();
  const typed = await citeTexts();
  const group = last && Object.values(last.store.citations).find((items) => items.some((x) => x.locator));
  check("marker group", typed.length === before + 1 && /^\[\d+[,\u2013-]\d+\]$/.test(typed[typed.length - 1]) && last.converted === 1
    && JSON.stringify(group) === JSON.stringify([{ id: "lee2010" }, { id: "mueller2020", locator: "12", label: "page" }]), { typed, group, converted: last?.converted });
  const text = await word(async (c) => { c.document.body.load("text"); await c.sync(); return c.document.body.text; });
  check("unknown key and e-mail stay as text", /Unknown: nokey2000/.test($("status").textContent) && text.includes("@nokey2000") && text.includes("anna@example.org") && !text.includes("@lee2010"), { status: $("status").textContent, text: text.slice(-200) });
}

const PLACEHOLDER = "[Your text]";
const HEADINGS = [
  ["1 Background", 1], ["1.1 Earlier work", 2], ["References for section 1", 2],
  ["2 Methods", 1], ["References for section 2", 2],
];

/** A document like a template: headings, each with a placeholder paragraph to fill in. */
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
      p.styleBuiltIn = Word.Style.normal;
      // The last placeholder is a list item, as templates often have them.
      if (i === HEADINGS.length - 1) p.startNewList();
    });
    await c.sync();
  });
}

/** Selects the first placeholder below a heading. */
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
    const i = where.findIndex((w) => w.value === "After" || w.value === "AdjacentAfter");
    ps.items[i].select();
    await c.sync();
  });
}

/** Types a text into the placeholder below a heading and puts the cursor at its end. */
async function fill(heading, text) {
  await selectPlaceholderAfter(heading);
  await word(async (c) => {
    c.document.getSelection().insertText(text, "Replace").select("End");
    await c.sync();
  });
}

// A template with two sections: citations in both, each bibliography in the
// placeholder under its "References for section" heading.
async function templateScenario() {
  await setupTemplate();
  const dt = new DataTransfer();
  dt.items.add(new File([BIB], "library.bib", { type: "text/plain" }));
  $("file").files = dt.files;
  $("file").dispatchEvent(new Event("change"));
  await idle();
  check("template import", /4 new/.test($("status").textContent), $("status").textContent);

  await fill("1 Background", "Tree rings record summer temperature.");
  await pick("mueller");
  await fill("1.1 Earlier work", "A survey mapped the glaciers.");
  await pick("survey2023");
  await selectPlaceholderAfter("References for section 1");
  $("bib").click();
  await idle();

  await fill("2 Methods", "The method is standard.");
  $("q").value = "lee"; $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", { ctrlKey: true });
  await pick("mueller");
  // A bibliography inserted with tracking on is derived content and stays untracked.
  await word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); });
  await selectPlaceholderAfter("References for section 2");
  $("bib").click();
  await idle();
  const tracked = await word(async (c) => {
    c.document.load("changeTrackingMode");
    const tcs = c.document.body.getTrackedChanges();
    tcs.load("items/type,items/text");
    await c.sync();
    const r = { mode: c.document.changeTrackingMode, changes: tcs.items.map((t) => [t.type, t.text]) };
    c.document.changeTrackingMode = "Off";
    await c.sync();
    return r;
  });
  check("template bibliography untracked", tracked.mode === "TrackAll" && tracked.changes.length === 0, tracked);

  const cites = await citeTexts();
  const bibs = await bibTexts();
  check("template sections", JSON.stringify(cites) === JSON.stringify(["[1]", "[2]", "[1,2]"]) && bibs.length === 2 && bibs[0].length === 2 && bibs[1].length === 2 && /^1\.\s*Lee K/.test(bibs[1][0]), { cites, bibs });
  const styles = await word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag");
    await c.sync();
    const ps = all.items.filter((x) => x.tag === BIB_TAG).map((x) => { const p = x.paragraphs; p.load("items/style,items/styleBuiltIn,items/isListItem"); return p; });
    await c.sync();
    return ps.flatMap((p) => p.items.map((q) => [q.style, q.styleBuiltIn, q.isListItem]));
  });
  check("template bibliography style", styles.length === 4 && styles.every(([, b, li]) => !/^Heading/.test(b) && !li), styles);
}

// The template scenario's document, saved and opened again: the pane's first
// refresh must find everything current.
async function verifyScenario() {
  check("reopened: first refresh writes nothing", first && first.citations.length === 0 && first.bibs.length === 0 && first.removes.length === 0, first && { c: first.citations, b: first.bibs.length, r: first.removes });
  check("reopened: sources and citations", first && first.cited === 3 && Object.keys(first.store.sources).length === 4, first && { cited: first.cited, unknown: first.unknown });
  // Before the sentence's period, so ahead of the citation behind it.
  await cursorAfter("The method is standard");
  await pick("snow");
  const bibs = await bibTexts();
  check("reopened: cite again", JSON.stringify(await citeTexts()) === JSON.stringify(["[1]", "[2]", "[1]", "[2,3]"]) && bibs[1].length === 3 && /^1\.\s*Alpine Survey Office\. Snow depth data/.test(bibs[1][0]), { cites: await citeTexts(), bibs });
}

// A document in the old stored format, built from the unit tests' fixture:
// one refresh upgrades the store, the tags and the typed markers.
async function legacyScenario() {
  const oldNs = /xmlns="([^"]+)"/.exec(legacyXml)[1];
  const [one, two, oldBib, foreign] = legacyDoc.controls;
  await word(async (c) => {
    const b = c.document.body;
    b.clear();
    c.document.changeTrackingMode = "Off";
    c.document.customXmlParts.add(legacyXml);
    const control = (range, { tag, title }) => { const cc = range.insertContentControl(); cc.tag = tag; cc.title = title; };
    const p1 = b.insertParagraph("First claim.", "Start");
    control(p1.insertText(one.text, "End"), one);
    const p2 = b.insertParagraph("Second claim.", "End");
    control(p2.insertText(two.text, "End"), two);
    for (const text of legacyDoc.paragraphs) b.insertParagraph(text, "End");
    control(b.insertParagraph("An old bibliography.", "End"), oldBib);
    control(b.insertParagraph(foreign.text, "End"), foreign);
    await c.sync();
  });
  $("refresh").click();
  await idle();

  const state = await word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag,items/title");
    const now = c.document.customXmlParts.getByNamespace(NS);
    const old = c.document.customXmlParts.getByNamespace(oldNs);
    now.load("items");
    old.load("items");
    c.document.body.load("text");
    await c.sync();
    return { tags: all.items.map((x) => x.tag), titles: all.items.map((x) => x.title), now: now.items.length, old: old.items.length, text: c.document.body.text };
  });
  const oldTags = [one.tag, two.tag, oldBib.tag];
  const cites = state.tags.filter((t) => t.startsWith(CITE_PREFIX)).length;
  check("legacy: store moved to the current part", state.now === 1 && state.old === 0 && last && Object.keys(last.store.sources).length > 0 && last.store.settings.locale === "de-DE", { now: state.now, old: state.old, settings: last?.store.settings });
  check("legacy: tags and titles renamed", !state.tags.some((t) => oldTags.includes(t)) && cites >= 2 && state.tags.filter((t) => t === BIB_TAG).length === 1
    && state.tags.includes(foreign.tag) && state.titles.filter((t) => t.startsWith("Cite: ")).length === cites, state);
  const texts = await citeTexts();
  check("legacy: old citations and old markers are cited", last?.converted > 0 && texts.length === cites && texts.every((t) => /^\[\d+([,\u2013-]\d+)*\]$/.test(t)), { texts, converted: last?.converted, status: $("status").textContent });
  check("legacy: no old marker is left", !/\[@[A-Za-z0-9]/.test(state.text), state.text);
  const bibs = await bibTexts();
  check("legacy: one bibliography with entries", bibs.length === 1 && bibs[0].length > 0 && /^1\.\s/.test(bibs[0][0]), bibs);

  $("refresh").click();
  await idle();
  check("legacy: a second refresh writes nothing", last && last.citations.length === 0 && last.bibs.length === 0 && !last.converted, last && { c: last.citations.length, b: last.bibs.length, converted: last.converted });
}

/** The runner names the scenario in the document's text. A document with citations was saved by an earlier one. */
async function mode() {
  return word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag");
    c.document.body.load("text");
    await c.sync();
    if (all.items.some(isCite)) return "verify";
    return /e2e:(template|legacy)/.exec(c.document.body.text)?.[1] ?? "main";
  });
}

document.addEventListener("siglum:ready", async () => {
  const m = await mode();
  post({ stage: "ready", mode: m });
  try {
    await { main: scenario, template: templateScenario, verify: verifyScenario, legacy: legacyScenario }[m]();
  } catch (e) {
    check("scenario", false, String(e && (e.stack || e.message || e)));
  }
  const final = await controls().catch((e) => String(e));
  post({ done: true, ok: results.every((r) => r.ok), failed: results.filter((r) => !r.ok).map((r) => r.name), controls: final });
});
window.addEventListener("error", (e) => post({ error: String(e.message) }));
window.addEventListener("unhandledrejection", (e) => post({ error: String(e.reason && (e.reason.stack || e.reason)) }));
