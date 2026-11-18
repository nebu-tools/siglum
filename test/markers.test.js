import { test } from "node:test";
import assert from "node:assert/strict";
import { findMarkers, parseSupplement } from "../src/model.js";

const keys = (text) => findMarkers(text).map((g) => g.items.map((x) => x.id));
const texts = (text) => findMarkers(text).map((g) => g.text);

test("a marker is @ and a key", () => {
  assert.deepEqual(findMarkers("As shown @smith2020 here"), [
    { start: 9, end: 19, text: "@smith2020", items: [{ id: "smith2020" }] },
  ]);
  assert.deepEqual(keys("@a at the start, at the end @b"), [["a"], ["b"]]);
  assert.deepEqual(keys("no marker here"), []);
  assert.deepEqual(keys(""), []);
});

test("keys follow Typst's label rules", () => {
  assert.deepEqual(keys("@fig:tree_rings-2 @DIN-1505.2 @müller2020 @2020"), [["fig:tree_rings-2", "DIN-1505.2", "müller2020", "2020"]]);
  assert.deepEqual(keys("@ alone, @. and @: are text"), []);
});

test("punctuation after a key stays in the sentence", () => {
  assert.deepEqual(texts("see @smith2020."), ["@smith2020"]);
  assert.deepEqual(texts("see @smith2020: it works"), ["@smith2020"]);
  assert.deepEqual(texts("(see @smith2020), then @lee2010; and @a.b.c..."), ["@smith2020", "@lee2010", "@a.b.c"]);
  assert.deepEqual(texts("@smith2020, @lee2010"), ["@smith2020", "@lee2010"]);
  assert.deepEqual(texts("is it @smith2020? yes, @lee2010!"), ["@smith2020", "@lee2010"]);
});

test("an @ inside a word is never a marker", () => {
  assert.deepEqual(keys("write to anna.mueller@example.org today"), []);
  assert.deepEqual(keys("anna@example.org, or cite @smith2020"), [["smith2020"]]);
  assert.deepEqual(keys("https://example.org/@anna and user@host:22"), []);
  assert.deepEqual(keys("3@4 a@b @c@d"), []);
  assert.deepEqual(keys("x.@a y_@b"), []);
});

test("what may come before a marker", () => {
  assert.deepEqual(keys("(@a) {@b} “@c” „@d“ – @e —@f '@g' «@h»"), [["a"], ["b"], ["c"], ["d"], ["e"], ["f"], ["g"], ["h"]]);
  assert.deepEqual(keys("a @b,@c;@d\t@e x"), [["b"], ["c"], ["d"], ["e"]]);
  // Anything else before the "@" makes it text.
  assert.deepEqual(keys("x@a /@b <@d> #@e *@f ]@g"), []);
});

test("keys are compared in composed form, and a mark belongs to its key", () => {
  // "e" and a combining acute accent are one character to the reader.
  assert.deepEqual(findMarkers("see @café and @é."), [
    { start: 4, end: 10, text: "@café", items: [{ id: "café" }] },
    { start: 15, end: 18, text: "@é", items: [{ id: "é" }] },
  ]);
  assert.deepEqual(keys("@müller @müller"), [["müller"]]);
  // A marker never ends inside a character: not before a joiner, not before a mark on its closing bracket.
  assert.deepEqual(keys("@ab‍"), []);
  assert.deepEqual(texts("@a[x]́"), ["@a"]);
  assert.deepEqual(keys("@́a"), []);
});

test("a supplement in brackets belongs to its key", () => {
  assert.deepEqual(findMarkers("@a[p. 12]")[0].items, [{ id: "a", locator: "12", label: "page" }]);
  assert.deepEqual(findMarkers("@a[see the table]")[0].items, [{ id: "a", suffix: "see the table" }]);
  assert.deepEqual(findMarkers("@a[]")[0], { start: 0, end: 4, text: "@a[]", items: [{ id: "a" }] });
  // A space ends the marker, and so does a bracket that never closes.
  assert.deepEqual(texts("@a [p. 12] and @b[p. 3"), ["@a", "@b"]);
  assert.deepEqual(texts("@a[unfinished"), ["@a"]);
  assert.deepEqual(texts("@a[x[y] and more"), ["@a"]);
  // So do a tab and a line break inside the brackets.
  assert.deepEqual(texts("@a[p.\t12] @b[p.\v3]"), ["@a", "@b"]);
});

test("brackets inside a supplement count in pairs", () => {
  assert.deepEqual(findMarkers("@a[outer[inner]] text")[0], { start: 0, end: 16, text: "@a[outer[inner]]", items: [{ id: "a", suffix: "outer[inner]" }] });
  assert.deepEqual(texts("@a[see [1] and [2]] @b"), ["@a[see [1] and [2]] @b"]);
  assert.deepEqual(texts("@a[x]] y"), ["@a[x]"]);
});

test("markers with only spaces between them are one citation", () => {
  const [g, ...rest] = findMarkers("Known @a @b[p. 3]  @c @d, and @e");
  assert.equal(g.text, "@a @b[p. 3]  @c @d");
  assert.deepEqual(g.items, [{ id: "a" }, { id: "b", locator: "3", label: "page" }, { id: "c" }, { id: "d" }]);
  assert.deepEqual(rest.map((x) => x.text), ["@e"]);
  assert.deepEqual(texts("@a\t@b and @c\n@d"), ["@a", "@b", "@c", "@d"]);
  assert.deepEqual(texts("@a\u00a0@b"), ["@a\u00a0@b"]);
  // A group ends before a word that is no marker.
  assert.deepEqual(texts("@a @b anna@example.org"), ["@a @b"]);
  assert.deepEqual(texts("@a @b@c"), ["@a"]);
});

test("a source repeated in a group counts once", () => {
  assert.deepEqual(findMarkers("@a @b @a")[0].items, [{ id: "a" }, { id: "b" }]);
  assert.deepEqual(findMarkers("@a @a[p. 2]")[0].items, [{ id: "a" }, { id: "a", locator: "2", label: "page" }]);
});

test("locators in English and German", () => {
  const loc = (t) => parseSupplement(t);
  assert.deepEqual(loc("p. 12"), { locator: "12", label: "page" });
  assert.deepEqual(loc("pp. 12-15"), { locator: "12-15", label: "page" });
  assert.deepEqual(loc("pp. 12–15"), { locator: "12-15", label: "page" });
  assert.deepEqual(loc("page 7"), { locator: "7", label: "page" });
  assert.deepEqual(loc("ch. 3"), { locator: "3", label: "chapter" });
  assert.deepEqual(loc("chapter 3"), { locator: "3", label: "chapter" });
  assert.deepEqual(loc("sec. 2"), { locator: "2", label: "section" });
  assert.deepEqual(loc("§ 4"), { locator: "4", label: "section" });
  assert.deepEqual(loc("fig. 2a"), { locator: "2a", label: "figure" });
  assert.deepEqual(loc("vol. 2"), { locator: "2", label: "volume" });
  assert.deepEqual(loc("S. 12"), { locator: "12", label: "page" });
  assert.deepEqual(loc("S. 12 – 14"), { locator: "12-14", label: "page" });
  assert.deepEqual(loc("Seite 12"), { locator: "12", label: "page" });
  assert.deepEqual(loc("Kap. 3"), { locator: "3", label: "chapter" });
  assert.deepEqual(loc("Abschnitt 2"), { locator: "2", label: "section" });
  assert.deepEqual(loc("Abb. 4"), { locator: "4", label: "figure" });
  assert.deepEqual(loc("p. xii"), { locator: "xii", label: "page" });
  assert.deepEqual(loc("pp. 12, 15"), { locator: "12, 15", label: "page" });
  assert.deepEqual(loc("S. 12 ff."), { locator: "12 ff.", label: "page" });
});

test("a bare number is a page", () => {
  assert.deepEqual(parseSupplement("12"), { locator: "12", label: "page" });
  assert.deepEqual(parseSupplement(" 12-15 "), { locator: "12-15", label: "page" });
  assert.deepEqual(parseSupplement("12f"), { locator: "12f", label: "page" });
});

test("any other supplement is a suffix", () => {
  assert.deepEqual(parseSupplement("see the table"), { suffix: "see the table" });
  assert.deepEqual(parseSupplement("p. 12 and elsewhere"), { suffix: "p. 12 and elsewhere" });
  assert.deepEqual(parseSupplement("mild"), { suffix: "mild" });
  assert.deepEqual(parseSupplement("xii"), { suffix: "xii" });
  assert.deepEqual(parseSupplement("table 3"), { suffix: "table 3" });
  assert.deepEqual(parseSupplement("p."), { suffix: "p." });
  assert.deepEqual(parseSupplement(""), {});
  assert.deepEqual(parseSupplement("   "), {});
  assert.deepEqual(parseSupplement(undefined), {});
});
