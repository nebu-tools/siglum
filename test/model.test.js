import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { createRenderer, htmlToRuns, resolveLocale, runsText } from "../src/render.js";
import { plan, storeToXml, storeFromXml, emptyStore, citeTitle, keysFromTitle, BIB_EMPTY, CITE_NEW, CITE_PREFIX, NS } from "../src/model.js";

const csl = (f) => fs.readFileSync(new URL("../csl/" + f, import.meta.url), "utf8");
const renderer = createRenderer({
  style: csl("nlm-citation-sequence-brackets.csl"),
  locales: { "en-US": csl("locales-en-US.xml"), "de-DE": csl("locales-de-DE.xml") },
});

const sources = {
  a: { id: "a", type: "article-journal", title: "Tree rings of <i>Pinus cembra</i>", author: [{ family: "Müller", given: "Anna" }], "container-title": "J Ecol", issued: { "date-parts": [[2020]] }, volume: "108", page: "301-310" },
  b: { id: "b", type: "report", title: "Glacier inventory AS-123", author: [{ literal: "Alpine Survey Office" }], issued: { "date-parts": [[2023]] } },
  c: { id: "c", type: "book", title: "Statistics", author: [{ family: "Lee", given: "K" }], issued: { "date-parts": [[2010]] }, publisher: "Springer", "publisher-place": "Berlin" },
  d: { id: "d", type: "webpage", title: "Snow depth data", URL: "https://example.org/snow", issued: { "date-parts": [[2024]] }, accessed: { "date-parts": [[2026, 11, 2]] } },
};
const store = (citations) => ({ sources, citations, settings: {} });
const cite = (id, current = "[?]", extra = {}) => ({ kind: "cite", tag: CITE_PREFIX + id, title: "", current, original: current, ...extra });
const bib = (current = "") => ({ kind: "bib", current });

test("htmlToRuns keeps italics, superscript and entities", () => {
  assert.deepEqual(htmlToRuns("Rings <i>of &#38; in</i><sup>2</sup> &#60;3"), [
    { text: "Rings " }, { text: "of & in", italic: true }, { text: "2", superscript: true }, { text: " <3" },
  ]);
});

test("each bibliography numbers its own section from 1", () => {
  const controls = [cite("1"), cite("2"), bib(), cite("3"), cite("4"), bib()];
  const p = plan(controls, store({ 1: [{ id: "b" }], 2: [{ id: "a" }, { id: "b" }], 3: [{ id: "c" }], 4: [{ id: "a" }] }), renderer);
  assert.deepEqual(p.citations.map((c) => [c.index, runsText(c.runs)]), [[0, "[1]"], [1, "[1,2]"], [3, "[1]"], [4, "[2]"]]);
  assert.equal(p.bibs.length, 2);
  const lists = p.bibs.map((b) => b.entries.map(runsText));
  assert.match(lists[0][0], /^1\.\tAlpine Survey Office\. Glacier inventory AS-123/);
  assert.match(lists[0][1], /^2\.\tMüller A\. Tree rings of Pinus cembra\. J Ecol\. 2020;108:301–10\.$/);
  assert.deepEqual(p.bibs[0].entries[1].find((r) => r.italic), { text: "Pinus cembra", italic: true });
  assert.match(lists[1][0], /^1\.\tLee K\. Statistics\. Berlin: Springer; 2010/);
  assert.equal(p.sections, 2);
  assert.equal(p.cited, 4);
});

test("citations after the last bibliography join it; without one all are numbered", () => {
  const p = plan([cite("1"), bib(), cite("2")], store({ 1: [{ id: "a" }], 2: [{ id: "b" }] }), renderer);
  assert.deepEqual(p.citations.map((c) => runsText(c.runs)), ["[1]", "[2]"]);
  assert.equal(p.bibs[0].entries.length, 2);
  const q = plan([cite("1"), cite("2")], store({ 1: [{ id: "a" }], 2: [{ id: "a" }] }), renderer);
  assert.deepEqual(q.citations.map((c) => runsText(c.runs)), ["[1]", "[1]"]);
  assert.deepEqual(q.bibs, []);
});

test("nothing to write when the document is current", () => {
  const first = plan([cite("1"), bib()], store({ 1: [{ id: "c" }] }), renderer);
  const text = first.bibs[0].entries.map(runsText).join("\r");
  const again = plan([cite("1", "[1]"), bib(text)], store({ 1: [{ id: "c" }] }), renderer);
  assert.deepEqual([again.citations, again.bibs, again.removes], [[], [], []]);
});

test("an empty bibliography says so, in the caller's words", () => {
  assert.deepEqual(plan([bib("old")], store({}), renderer).bibs[0].entries, [[{ text: BIB_EMPTY }]]);
  assert.deepEqual(plan([bib("old")], store({}), renderer, { bibEmpty: "Nothing yet." }).bibs[0].entries, [[{ text: "Nothing yet." }]]);
  assert.deepEqual(plan([bib("Nothing yet.")], store({}), renderer, { bibEmpty: "Nothing yet." }).bibs, []);
});

test("the locale sets the language of the entries", () => {
  const entry = (locale) => runsText(plan([cite("1"), bib()], store({ 1: [{ id: "d" }] }), renderer, { locale }).bibs[0].entries[0]);
  assert.match(entry("en-US"), /\[cited 2026 Nov 2\]\. Available from: https:\/\/example\.org\/snow$/);
  assert.match(entry("de-DE"), /\[zitiert 2\. November 2026\]\. Verfügbar unter: https:\/\/example\.org\/snow$/);
  assert.equal(entry("fr-FR"), entry("en-US"));
});

test("a locale tag resolves to a bundled locale", () => {
  const have = ["en-US", "de-DE"];
  assert.equal(resolveLocale("de-DE", have), "de-DE");
  assert.equal(resolveLocale("de-AT", have), "de-DE");
  assert.equal(resolveLocale("de", have), "de-DE");
  assert.equal(resolveLocale("en_GB", have), "en-US");
  assert.equal(resolveLocale("ja-JP", have), "en-US");
  assert.equal(resolveLocale(undefined, have), "en-US");
  assert.deepEqual(renderer.locales, have);
});

test("a supplement that is no locator follows the number", () => {
  const p = plan([cite("1"), cite("2")], store({ 1: [{ id: "a", suffix: "see the table" }], 2: [{ id: "a", locator: "12", label: "page" }, { id: "b" }] }), renderer);
  assert.deepEqual(p.citations.map((c) => runsText(c.runs)), ["[1, see the table]", "[1,2]"]);
});

test("emptied citations go away, tracked deletions wait", () => {
  const p = plan(
    [cite("1", ""), cite("2", " ", { original: "" }), cite("3", "", { original: "[1]" }), cite("4", CITE_NEW), bib()],
    store({ 1: [{ id: "a" }], 2: [{ id: "a" }], 3: [{ id: "a" }], 4: [{ id: "b" }] }),
    renderer,
  );
  assert.deepEqual(p.removes, [0, 1]);
  assert.equal(p.pending, 1);
  assert.deepEqual(p.citations.map((c) => [c.index, runsText(c.runs)]), [[3, "[1]"]]);
  assert.equal(p.bibs[0].entries.length, 1);
});

test("a citation pasted from another document is recovered from its title", () => {
  const p = plan([{ ...cite("x9", "[7]"), title: citeTitle([{ id: "b" }, { id: "a" }]) }], store({}), renderer);
  assert.deepEqual(p.store.citations.x9, [{ id: "b" }, { id: "a" }]);
  assert.equal(runsText(p.citations[0].runs), "[1,2]");
  const lost = plan([{ ...cite("y", "[7]"), title: citeTitle([{ id: "zz" }]) }], store({}), renderer);
  assert.deepEqual(lost.unknown, [{ index: 0, keys: ["zz"], text: "[7]" }]);
  assert.deepEqual(lost.citations, []);
});

test("titles round-trip keys unless cut short", () => {
  assert.equal(citeTitle([{ id: "a" }, { id: "b2" }]), "Cite: a; b2");
  assert.deepEqual(keysFromTitle(citeTitle([{ id: "a" }, { id: "b2" }])), ["a", "b2"]);
  const long = citeTitle(Array.from({ length: 12 }, (_, i) => ({ id: "mueller202" + i })));
  assert.equal(long.length, 64);
  assert.equal(keysFromTitle(long), null);
});

test("the store survives XML", () => {
  const s = { sources: { a: { id: "a", title: "A < B & C > \"D\"" } }, citations: { z: [{ id: "a" }] }, settings: { locale: "de-DE" } };
  const xml = storeToXml(s);
  assert.equal(NS, "urn:siglum:store:1");
  assert.match(xml, /^<\?xml[^>]*\?><siglum xmlns="urn:siglum:store:1">\{"sources":/);
  assert.deepEqual(storeFromXml(xml), s);
  assert.deepEqual(storeFromXml(xml.replace("<?xml version=\"1.0\" encoding=\"UTF-8\"?>", "")), s);
  assert.throws(() => storeFromXml(""));
  assert.deepEqual(storeFromXml(storeToXml(emptyStore("de-DE"))), { sources: {}, citations: {}, settings: { locale: "de-DE" } });
});
