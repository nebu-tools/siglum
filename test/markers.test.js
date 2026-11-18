import { test } from "node:test";
import assert from "node:assert/strict";
import { findMarkers, parseSupplement } from "../src/model.js";

const keys = (text) => findMarkers(text).map((g) => g.items.map((x) => x.id));
const texts = (text) => findMarkers(text).map((g) => g.text);

test("a marker is @ and a key", () => {
  assert.deepEqual(findMarkers("As shown @smith2020 here"), [
    { start: 9, end: 19, text: "@smith2020", items: [{ id: "smith2020" }], nth: 0, of: 1 },
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

test("quotes, brackets and dashes may come before a marker", () => {
  assert.deepEqual(keys("(@a) [@b] “@c” „@d“ – @e —@f"), [["a"], ["b"], ["c"], ["d"], ["e"], ["f"]]);
  assert.deepEqual(keys("a @b"), [["b"]]);
});

test("a supplement in brackets belongs to its key", () => {
  assert.deepEqual(findMarkers("@a[p. 12]")[0].items, [{ id: "a", locator: "12", label: "page" }]);
  assert.deepEqual(findMarkers("@a[see the table]")[0].items, [{ id: "a", suffix: "see the table" }]);
  assert.deepEqual(findMarkers("@a[]")[0], { start: 0, end: 4, text: "@a[]", items: [{ id: "a" }], nth: 0, of: 1 });
  // A space ends the marker, and so does a bracket that never closes.
  assert.deepEqual(texts("@a [p. 12] and @b[p. 3"), ["@a", "@b"]);
  assert.deepEqual(texts("@a[x[y]]"), ["@a"]);
});

test("markers with only spaces between them are one citation", () => {
  const [g, ...rest] = findMarkers("Known @a @b[p. 3]  @c @d, and @e");
  assert.equal(g.text, "@a @b[p. 3]  @c @d");
  assert.deepEqual(g.items, [{ id: "a" }, { id: "b", locator: "3", label: "page" }, { id: "c" }, { id: "d" }]);
  assert.deepEqual(rest.map((x) => x.text), ["@e"]);
  assert.deepEqual(texts("@a\t@b and @c\n@d"), ["@a", "@b", "@c", "@d"]);
  // A group ends before a word that is no marker.
  assert.deepEqual(texts("@a @b anna@example.org"), ["@a @b"]);
  assert.deepEqual(texts("@a @b@c"), ["@a"]);
});

test("a source repeated in a group counts once", () => {
  assert.deepEqual(findMarkers("@a @b @a")[0].items, [{ id: "a" }, { id: "b" }]);
  assert.deepEqual(findMarkers("@a @a[p. 2]")[0].items, [{ id: "a" }, { id: "a", locator: "2", label: "page" }]);
});

test("nth and of place a marker among equal texts in its paragraph", () => {
  const g = findMarkers("First @a, then @ab, then @a and anna@a.example @a");
  assert.deepEqual(g.map((x) => [x.text, x.nth, x.of]), [["@a", 0, 5], ["@ab", 0, 1], ["@a", 2, 5], ["@a", 4, 5]]);
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
