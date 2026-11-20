import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { createRenderer, runsText } from "../src/render.js";

const csl = (f) => fs.readFileSync(new URL("../csl/" + f, import.meta.url), "utf8");
const renderer = createRenderer({
  style: csl("nlm-citation-sequence-brackets.csl"),
  locales: { "en-US": csl("locales-en-US.xml"), "de-DE": csl("locales-de-DE.xml") },
});
const lib = {
  a: { id: "a", type: "article-journal", title: "One", author: [{ family: "Alpha", given: "A" }], issued: { "date-parts": [[2020]] }, "container-title": "J", volume: "1", page: "1-2" },
  b: { id: "b", type: "article-journal", title: "Two", author: [{ family: "Beta", given: "B" }], issued: { "date-parts": [[2021]] }, "container-title": "J", volume: "2", page: "3-4" },
  c: { id: "c", type: "article-journal", title: "Three", author: [{ family: "Gamma", given: "C" }], issued: { "date-parts": [[2022]] }, "container-title": "J", volume: "3", page: "5-6" },
};
const text = (items, locale = "en-US") => runsText(renderer.render([{ citations: [{ items }] }], lib, locale)[0].citations[0]);

test("a locator prints after the number", () => {
  assert.equal(text([{ id: "a", locator: "12", label: "page" }]), "[1, p. 12]");
});

test("locator labels follow the citation language", () => {
  assert.equal(text([{ id: "a", locator: "12", label: "page" }], "de-DE"), "[1, S. 12]");
  assert.equal(text([{ id: "a", locator: "3", label: "chapter" }]), "[1, chap. 3]");
  assert.equal(text([{ id: "a", locator: "12-15", label: "page" }]), "[1, pp. 12–5]");
});

test("a locator breaks the range collapse around it", () => {
  assert.equal(text([{ id: "a" }, { id: "b" }, { id: "c" }]), "[1–3]");
  assert.equal(text([{ id: "a", locator: "12", label: "page" }, { id: "b" }, { id: "c" }]), "[1, p. 12,2,3]");
});

test("a prefix prints before, a suffix after", () => {
  assert.equal(text([{ id: "a", prefix: "see " }]), "[see 1]");
  assert.equal(text([{ id: "a", suffix: "emphasis added" }]), "[1, emphasis added]");
  assert.equal(text([{ id: "a", locator: "12", label: "page", suffix: "emphasis added" }]), "[1, p. 12, emphasis added]");
  assert.equal(text([{ id: "a", locator: "12", label: "page", prefix: "see " }]), "[see 1, p. 12]");
});

test("suppress-author is inert in the numbered style", () => {
  assert.equal(text([{ id: "a", suppressAuthor: true }]), "[1]");
});

test("the bundled style emits no small caps runs", () => {
  // Per-control writes cannot render small caps (Word's Font has no such
  // property), so a style Siglum bundles must not produce any; author-date
  // styles with small caps stay out until writes go through OOXML.
  const combos = [[{ id: "a" }], [{ id: "a", locator: "12", label: "page" }], [{ id: "a" }, { id: "b" }], [{ id: "a", prefix: "see " }], [{ id: "a", suffix: "x" }]];
  for (const items of combos) {
    for (const run of renderer.render([{ citations: [{ items }] }], lib, "en-US")[0].citations[0]) {
      assert.ok(!run.smallCaps, JSON.stringify(run));
    }
  }
});
