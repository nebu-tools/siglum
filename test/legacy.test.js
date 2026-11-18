import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { upgradeLegacy, storeFromLegacyXml, upgradeTag, upgradeTitle, upgradeMarkers } from "../src/legacy.js";
import { plan, storeToXml, storeFromXml, findMarkers, keysFromTitle, BIB_TAG, CITE_PREFIX, NS } from "../src/model.js";
import { createRenderer, runsText } from "../src/render.js";
import { documentText } from "../src/messages.js";
import { fakeDocument } from "./fake-word.js";

// A document as the version before the rename left it.
const read = (f) => fs.readFileSync(new URL(f, import.meta.url), "utf8");
const OLD_XML = read("fixtures/legacy/store.xml");
const OLD_DOC = JSON.parse(read("fixtures/legacy/document.json"));
const csl = (f) => read("../csl/" + f);
const renderer = createRenderer({
  style: csl("nlm-citation-sequence-brackets.csl"),
  locales: { "en-US": csl("locales-en-US.xml"), "de-DE": csl("locales-de-DE.xml") },
});

/** Runs the upgrade the way word.js does: the store is saved before the old part goes. */
async function upgrade(doc, context) {
  const order = [];
  await upgradeLegacy(context, {
    adopt: async (store) => {
      order.push({ adopt: Object.keys(store.sources).length, oldParts: doc.parts.length });
      context.document.customXmlParts.add(storeToXml(store));
      await context.sync();
    },
    untracked: async (fn) => {
      const was = doc.tracking;
      doc.tracking = "Off";
      try { return await fn(); } finally { doc.tracking = was; }
    },
  });
  return order;
}

test("an old store reads into the current shape", () => {
  const store = storeFromLegacyXml(OLD_XML);
  assert.deepEqual(Object.keys(store), ["sources", "citations", "settings"]);
  assert.deepEqual(Object.keys(store.sources), ["mueller2020", "lee2010", "snow2024"]);
  assert.equal(store.sources.mueller2020.title, "Tree rings & summer temperature in the Alps");
  assert.deepEqual(store.citations.k3x9a1, [{ id: "mueller2020" }, { id: "lee2010", locator: "12", label: "page" }]);
  // The old version wrote German, so an old document stays German.
  assert.deepEqual(store.settings, { locale: "de-DE" });
  assert.equal(storeFromLegacyXml(""), null);
  assert.equal(storeFromLegacyXml(storeToXml(store)), null);
});

test("old tags and titles get the current names", () => {
  assert.equal(upgradeTag("zitat:k3x9a1"), CITE_PREFIX + "k3x9a1");
  assert.equal(upgradeTag("literaturverzeichnis"), BIB_TAG);
  assert.equal(upgradeTag(CITE_PREFIX + "k3x9a1"), null);
  assert.equal(upgradeTag("reviewer-note"), null);
  assert.equal(upgradeTag(undefined), null);
  assert.deepEqual(keysFromTitle(upgradeTitle("Zitat: mueller2020; lee2010")), ["mueller2020", "lee2010"]);
  assert.equal(upgradeTitle("Zitatfreie Notiz"), "Zitatfreie Notiz");
  assert.equal(upgradeTitle(undefined), undefined);
});

test("old markers are rewritten to the current syntax", () => {
  assert.equal(upgradeMarkers("[@a]"), "@a");
  assert.equal(upgradeMarkers("see [@a; @b2020, S. 12-14]."), "see @a @b2020[S. 12-14].");
  assert.equal(upgradeMarkers("[ @a ;@b ] and [@c, see ch. 2]"), "@a @b and @c[see ch. 2]");
  assert.deepEqual(findMarkers(upgradeMarkers("[@a; @b2020, S. 12-14]"))[0].items, [{ id: "a" }, { id: "b2020", locator: "12-14", label: "page" }]);
  // What was no marker before is none now.
  for (const text of ["[@a; b]", "[email]", "[@ home]", "[1]", "[@a.]", "@a", "[@a\r@b]"]) assert.equal(upgradeMarkers(text), text);
});

test("a document in the old format is upgraded in place", async () => {
  const { doc, context } = fakeDocument({ parts: [OLD_XML, "<other xmlns=\"urn:example:other\"/>"], controls: OLD_DOC.controls, paragraphs: OLD_DOC.paragraphs });
  const order = await upgrade(doc, context);

  // The old part went only after the new one was saved.
  assert.deepEqual(order, [{ adopt: 3, oldParts: 2 }]);
  assert.equal(doc.parts.length, 2);
  assert.ok(doc.parts.some((p) => p.xml.includes("urn:example:other")));
  const part = doc.parts.find((p) => p.xml.includes(NS));
  assert.deepEqual(storeFromXml(part.xml), storeFromLegacyXml(OLD_XML));

  assert.deepEqual(doc.controls.map((c) => [c.tag, c.title]), [
    [CITE_PREFIX + "k3x9a1", "Cite: mueller2020; lee2010"],
    [CITE_PREFIX + "p7q2b4", "Cite: snow2024"],
    [BIB_TAG, "Literaturverzeichnis"],
    ["reviewer-note", "Zitatfreie Notiz"],
  ]);

  assert.deepEqual(doc.paragraphs, OLD_DOC.upgraded);
  assert.equal(doc.log.length, 2);
  assert.ok(doc.log.every((e) => e.tracking === "Off"));
  assert.equal(doc.tracking, "TrackAll");
});

test("an upgraded document plans to the text it had", async () => {
  const { doc, context } = fakeDocument({ parts: [OLD_XML], controls: OLD_DOC.controls, paragraphs: OLD_DOC.paragraphs });
  await upgrade(doc, context);
  const store = storeFromXml(doc.parts[0].xml);
  const controls = doc.controls.filter((c) => c.tag.startsWith(CITE_PREFIX) || c.tag === BIB_TAG).map((c) => (c.tag === BIB_TAG
    ? { kind: "bib", current: c.text }
    : { kind: "cite", tag: c.tag, title: c.title, current: c.text, original: c.text }));
  const { locale } = store.settings;
  const p = plan(controls, store, renderer, { locale, bibEmpty: documentText(locale).bibEmpty });
  // The numbers in the fixture are still right, so no citation is rewritten.
  assert.deepEqual(p.citations, []);
  assert.deepEqual(p.unknown, []);
  assert.equal(p.cited, 2);
  const entries = p.bibs[0].entries.map(runsText);
  assert.equal(entries.length, 3);
  assert.match(entries[2], /^3\.\tAlpine Survey Office\. Snow depth data \[Internet\]\. 2024 \[zitiert 2\. November 2026\]\. Verfügbar unter: /);
});

test("the upgrade runs once and leaves a current document alone", async () => {
  const { doc, context } = fakeDocument({ parts: [OLD_XML], controls: OLD_DOC.controls, paragraphs: OLD_DOC.paragraphs });
  await upgrade(doc, context);
  const after = JSON.stringify(doc);
  const syncs = doc.syncs;
  assert.deepEqual(await upgrade(doc, context), []);
  // One sync to look, and one more because "[@ home]" still looks like an old marker.
  assert.equal(doc.syncs, syncs + 2);
  assert.equal(JSON.stringify({ ...doc, syncs }), after);

  const fresh = fakeDocument({ paragraphs: ["Write to anna@example.org or cite @lee2010 and [1]."] });
  assert.deepEqual(await upgrade(fresh.doc, fresh.context), []);
  assert.equal(fresh.doc.syncs, 1);
  assert.deepEqual(fresh.doc.paragraphs, ["Write to anna@example.org or cite @lee2010 and [1]."]);
});

test("only the legacy module and its fixtures know the old names", () => {
  const old = /zitat|newga|quellen|literaturverzeichnis/i;
  const root = new URL("../", import.meta.url);
  const allowed = new Set(["src/legacy.js", "src/messages.js", "test/legacy.test.js", "test/messages.test.js", "manifest.xml"]);
  const hits = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(new URL(dir, root), { withFileTypes: true })) {
      const rel = dir + e.name;
      if (e.isDirectory()) { if (rel !== "test/fixtures") walk(rel + "/"); continue; }
      if (/\.(js|mjs|html|css|xml|json|ps1|yml)$/.test(e.name) && !allowed.has(rel) && old.test(fs.readFileSync(new URL(rel, root), "utf8"))) hits.push(rel);
    }
  };
  for (const dir of ["src/", "test/", "e2e/", "scripts/", ".github/"]) walk(dir);
  if (old.test(fs.readFileSync(new URL("package.json", root), "utf8"))) hits.push("package.json");
  assert.deepEqual(hits, []);
  // German words in the translation are no stored names: none of them is a tag, a namespace or a JSON key.
  for (const f of ["src/messages.js", "manifest.xml"]) assert.doesNotMatch(read("../" + f), /newga|zitat:|"quellen"|"zitate"|"literaturverzeichnis"/);
});
