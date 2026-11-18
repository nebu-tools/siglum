import { test } from "node:test";
import assert from "node:assert/strict";
import { lookup, parseIdent, cleanCsl, fromEsummary } from "../src/lookup.js";

// A real esummary record, shortened.
const SUMMARY = {
  uid: "2231712", pubdate: "1990 Oct 5", source: "J Mol Biol", title: "Basic local alignment search tool.",
  volume: "215", issue: "3", pages: "403-10",
  authors: [{ name: "Altschul SF", authtype: "Author" }, { name: "Gish W", authtype: "Author" }, { name: "Example Study Group", authtype: "CollectiveName" }],
  articleids: [{ idtype: "pubmed", value: "2231712" }, { idtype: "doi", value: "10.1016/S0022-2836(05)80360-2" }],
};

test("DOIs and PMIDs in their usual spellings", () => {
  assert.deepEqual(parseIdent("https://doi.org/10.1038/171737a0."), { doi: "10.1038/171737a0" });
  assert.deepEqual(parseIdent("doi:10.1000/xyz(1)"), { doi: "10.1000/xyz(1)" });
  assert.deepEqual(parseIdent("PMID: 2231712"), { pmid: "2231712" });
  assert.equal(parseIdent("Müller 2020"), null);
});

test("PubMed summaries become CSL", () => {
  assert.deepEqual(fromEsummary(SUMMARY), {
    type: "article-journal", PMID: "2231712", title: "Basic local alignment search tool",
    "container-title": "J Mol Biol", volume: "215", issue: "3", page: "403-10", DOI: "10.1016/S0022-2836(05)80360-2",
    author: [{ family: "Altschul", given: "SF" }, { family: "Gish", given: "W" }, { literal: "Example Study Group" }],
    issued: { "date-parts": [[1990, 10, 5]] },
  });
});

test("Crossref CSL loses what CSL does not know", () => {
  assert.equal(cleanCsl({ type: "journal-article" }).type, "article-journal");
  const c = cleanCsl({ type: "article", title: ["T"], author: [{ family: "A", given: "B", affiliation: [], sequence: "first" }], reference: [{}], "container-title": "J", issued: { "date-parts": [[2020, 1]], timestamp: 1 } });
  assert.deepEqual(c, { type: "article-journal", title: "T", author: [{ family: "A", given: "B" }], "container-title": "J", issued: { "date-parts": [[2020, 1]] } });
});

test("lookup asks doi.org for CSL and E-utilities for PMIDs", async () => {
  const calls = [];
  const fake = async (url, opts) => {
    calls.push([url, opts?.headers?.Accept]);
    if (url.startsWith("https://doi.org/")) return { ok: true, json: async () => ({ type: "article-journal", title: "T", DOI: "10.1000/x" }) };
    return { ok: true, json: async () => ({ result: { 2231712: SUMMARY, 1: { uid: "1", error: "cannot get document summary" } } }) };
  };
  assert.equal((await lookup("10.1000/x", fake)).title, "T");
  assert.equal((await lookup("2231712", fake)).PMID, "2231712");
  await assert.rejects(lookup("1", fake), (e) => e.key === "error.pmidNotFound" && /^PMID not found \(cannot get/.test(e.message));
  await assert.rejects(lookup("no identifier", fake), (e) => e.key === "error.noIdentifier");
  await assert.rejects(lookup("10.1000/gone", async () => ({ ok: false, status: 404 })), (e) => e.key === "error.doiNotFound" && e.params.reason === 404);
  assert.deepEqual(calls[0], ["https://doi.org/10.1000/x", "application/vnd.citationstyles.csl+json"]);
  assert.match(calls[1][0], /^https:\/\/eutils\.ncbi\.nlm\.nih\.gov\/entrez\/eutils\/esummary\.fcgi\?db=pubmed&retmode=json&id=2231712$/);
  assert.equal(calls[1][1], undefined);
});
