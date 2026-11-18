import { test } from "node:test";
import assert from "node:assert/strict";
import { buildIndex, search } from "../src/search.js";

const lib = {
  mueller2020: { id: "mueller2020", title: "Tree rings and summer temperature in the Alps", author: [{ family: "Müller", given: "Anna" }, { family: "Smith", given: "John" }], "container-title": "J Ecol", issued: { "date-parts": [[2020]] }, DOI: "10.1000/JEcol.2001" },
  smith2019: { id: "smith2019", title: "Snow cover and spring phenology", author: [{ family: "Smith", given: "John" }], "container-title": "Glob Chang Biol", issued: { "date-parts": [[2019]] }, PMID: "31234567" },
  survey2023: { id: "survey2023", title: "Glacier inventory AS-123", author: [{ literal: "Alpine Survey Office" }], issued: { "date-parts": [[2023]] } },
  schmidt2021: { id: "schmidt2021", title: "Mueller matrices in polarised light", author: [{ family: "Schmidt", given: "Eva" }], issued: { "date-parts": [[2021]] } },
};
const idx = buildIndex(lib);
const ids = (q) => search(idx, q).map((x) => x.id);

test("an empty query lists everything, newest first", () => {
  assert.deepEqual(ids(""), ["survey2023", "schmidt2021", "mueller2020", "smith2019"]);
});

test("every word must match; first authors rank above co-authors", () => {
  assert.deepEqual(ids("smith"), ["smith2019", "mueller2020"]);
  assert.deepEqual(ids("smith 2020"), ["mueller2020"]);
  assert.deepEqual(ids("smith glob"), ["smith2019"]);
});

test("umlauts, prefixes and typos", () => {
  assert.equal(ids("müller")[0], "mueller2020");
  assert.equal(ids("muel")[0], "mueller2020");
  assert.equal(ids("mueler")[0], "mueller2020");
  assert.deepEqual(ids("temprature"), ["mueller2020"]);
});

test("an author outranks the same word in a title", () => {
  assert.deepEqual(ids("mueller"), ["mueller2020", "schmidt2021"]);
});

test("keys, institutions, DOI and PMID", () => {
  assert.equal(ids("survey2023")[0], "survey2023");
  assert.deepEqual(ids("alpine survey"), ["survey2023"]);
  assert.deepEqual(ids("10.1000/jecol.2001"), ["mueller2020"]);
  assert.deepEqual(ids("31234567"), ["smith2019"]);
  assert.deepEqual(ids("nothing like this"), []);
});
