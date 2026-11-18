import { test } from "node:test";
import assert from "node:assert/strict";
import { locateMarkers } from "../src/word.js";
import { BIB_TAG, CITE_PREFIX } from "../src/model.js";
import { fakeDocument } from "./fake-word.js";

/** The span of the `nth` occurrence of `needle` in paragraph `p`, as a control: {tag, start, end}. */
function around(paragraphs, p, needle, tag, nth = 0) {
  let at = -1;
  for (let i = 0; i <= nth; i++) at = paragraphs[p].indexOf(needle, at + 1);
  assert.notEqual(at, -1, needle);
  const start = paragraphs.slice(0, p).reduce((n, x) => n + x.length + 1, 0) + at;
  return { tag, start, end: start + needle.length };
}

async function locate(paragraphs, controls = []) {
  const { doc, context } = fakeDocument({ paragraphs, controls });
  const r = await locateMarkers(context);
  const text = paragraphs.join("\r");
  return {
    ...r,
    doc,
    // Each marker as the document's text at its range, and where it starts.
    at: r.markers.map((m) => [text.slice(m.range.start, m.range.end), m.range.start]),
  };
}

test("a document without an @ costs one sync", async () => {
  const r = await locate(["No markers here."]);
  assert.deepEqual(r.markers, []);
  assert.equal(r.doc.syncs, 1);
});

test("the same marker twice in a paragraph: two ranges from one search", async () => {
  const r = await locate(["First @a, then @a again."]);
  assert.deepEqual(r.at, [["@a", 6], ["@a", 15]]);
  assert.deepEqual(r.markers.map((m) => m.items), [[{ id: "a" }], [{ id: "a" }]]);
  assert.deepEqual(r.doc.searches, ["@", "@a"]);
  assert.deepEqual(r.lost, []);
});

test("a marker whose text begins another key's marker", async () => {
  const r = await locate(["See @ab and @a, then @abc."]);
  assert.deepEqual(r.at, [["@ab", 4], ["@a", 12], ["@abc", 21]]);
  assert.deepEqual(r.markers.map((m) => m.items[0].id), ["ab", "a", "abc"]);
});

test("a marker beside an e-mail address", async () => {
  const r = await locate(["Write to anna@a.org, or cite @a."]);
  assert.deepEqual(r.at, [["@a", 29]]);
});

test("markers in several paragraphs keep to their own", async () => {
  const r = await locate(["One @a.", "Two @a and @b.", "No marker."]);
  assert.deepEqual(r.at, [["@a", 4], ["@a", 12], ["@b", 19]]);
});

test("a group longer than a search is located by its start and end", async () => {
  const keys = Array.from({ length: 30 }, (_, i) => `source${String(i + 1).padStart(2, "0")}`);
  const group = keys.map((k) => "@" + k).join(" ");
  assert.ok(group.length > 255);
  const r = await locate([`Many sources ${group}, then more text.`]);
  assert.deepEqual(r.lost, []);
  assert.deepEqual(r.at, [[group, 13]]);
  assert.deepEqual(r.markers[0].items.map((x) => x.id), keys);
  assert.ok(r.doc.searches.every((s) => s.length <= 255));
});

test("a long group that occurs twice, and one with carets", async () => {
  const group = Array.from({ length: 40 }, (_, i) => `@k${i}[^${i}]`).join(" ");
  assert.ok(group.length > 255);
  const r = await locate([`${group} and again ${group}.`]);
  assert.deepEqual(r.lost, []);
  assert.deepEqual(r.at, [[group, 0], [group, group.length + 11]]);
  assert.deepEqual(r.markers[0].items[3], { id: "k3", suffix: "^3" });
});

test("a caret in a supplement is searched as ^^", async () => {
  const r = await locate(["See @a[x^2 term]."]);
  assert.deepEqual(r.at, [["@a[x^2 term]", 4]]);
  assert.deepEqual(r.doc.searches, ["@", "@a[x^^2 term]"]);
});

test("a marker inside a citation or a bibliography is part of its text", async () => {
  const paragraphs = ["Cited [1 @a] here, and @a outside.", "1. Lee K. Mail @a in an entry.", "A note @a in another add-in's control."];
  const cite = around(paragraphs, 0, "[1 @a]", CITE_PREFIX + "x1");
  const bib = around(paragraphs, 1, paragraphs[1], BIB_TAG);
  const foreign = around(paragraphs, 2, "note @a in", "other-addin");
  const r = await locate(paragraphs, [cite, bib, foreign]);
  assert.deepEqual(r.at, [["@a", 23], ["@a", paragraphs[0].length + 1 + paragraphs[1].length + 1 + 7]]);
});

test("a marker two levels deep inside a citation or a bibliography is part of its text", async () => {
  const paragraphs = ["Cited [1 @a] here, and @a outside.", "1. Lee K. Mail @a in an entry."];
  const cite = around(paragraphs, 0, "[1 @a]", CITE_PREFIX + "x1");
  const inCite = around(paragraphs, 0, "@a]", "other-addin");
  const bib = around(paragraphs, 1, paragraphs[1], BIB_TAG);
  const inBib = around(paragraphs, 1, "Mail @a in", "other-addin");
  const r = await locate(paragraphs, [cite, inCite, bib, inBib]);
  assert.deepEqual(r.at, [["@a", 23]]);
});

test("many equal markers in a paragraph cost one search and four syncs", async () => {
  const paragraph = Array.from({ length: 3000 }, () => "@a,").join(" ");
  const r = await locate([paragraph, paragraph]);
  assert.equal(r.markers.length, 6000);
  assert.deepEqual(r.doc.searches, ["@", "@a", "@a"]);
  assert.equal(r.doc.syncs, 4);
});
