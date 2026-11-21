import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { MESSAGES, LANGUAGES, pickLanguage, translator, documentText, UserError } from "../src/messages.js";

const read = (f) => fs.readFileSync(new URL("../src/" + f, import.meta.url), "utf8");
const params = (msg) => [...String(msg).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();

test("every language has every message", () => {
  const want = Object.keys(MESSAGES.en).sort();
  assert.deepEqual(Object.keys(MESSAGES).sort(), Object.keys(LANGUAGES).sort());
  for (const [lang, table] of Object.entries(MESSAGES)) {
    assert.deepEqual(Object.keys(table).sort(), want, lang);
    for (const [key, msg] of Object.entries(table)) {
      const forms = [msg].flat();
      assert.ok(forms.every((m) => typeof m === "string" && m.trim()), `${lang} ${key} is empty`);
      assert.equal(Array.isArray(msg), Array.isArray(MESSAGES.en[key]), `${lang} ${key}: plural form`);
      if (Array.isArray(msg)) assert.equal(msg.length, 2, `${lang} ${key}`);
      for (const m of forms) assert.deepEqual(params(m), params([MESSAGES.en[key]].flat()[0]), `${lang} ${key}: parameters`);
    }
  }
});

test("the pane and the code use only messages that exist", () => {
  const html = read("taskpane.html");
  const used = [...html.matchAll(/data-i18n(?:-[a-z-]+)?="([^"]+)"/g)].map((m) => m[1]);
  assert.ok(used.length > 10);
  for (const f of ["taskpane.js", "word.js", "ooxml.js", "lookup.js", "messages.js"]) {
    used.push(...[...read(f).matchAll(/\b(?:t|\w*Error)\(\s*"([^"]+)"\s*[,)]/g)].map((m) => m[1]));
  }
  for (const key of used) assert.ok(key in MESSAGES.en, key);
  // Keys that the code builds from parts.
  const built = (key) => /^(locale|style)\./.test(key) || ["results.none", "results.emptyLibrary", "lookup.added", "lookup.known"].includes(key);
  const unused = Object.keys(MESSAGES.en).filter((key) => !used.includes(key) && !built(key));
  assert.deepEqual(unused, []);
});

test("the pane's default text is the English message", () => {
  const html = read("taskpane.html");
  for (const m of html.matchAll(/data-i18n="([^"]+)">([^<]*)</g)) assert.equal(m[2], MESSAGES.en[m[1]], m[1]);
  for (const m of html.matchAll(/\b(placeholder|aria-label|title)="([^"]*)" data-i18n-\1="([^"]+)"/g)) assert.equal(m[2], MESSAGES.en[m[3]], m[3]);
});

test("the language follows the tag, English when there is no translation", () => {
  assert.equal(pickLanguage("de-DE"), "de");
  assert.equal(pickLanguage("de_AT"), "de");
  assert.equal(pickLanguage("DE"), "de");
  assert.equal(pickLanguage("en-GB"), "en");
  assert.equal(pickLanguage("fr-FR"), "en");
  assert.equal(pickLanguage(undefined), "en");
  assert.equal(pickLanguage("toString"), "en");
});

test("messages take parameters and plurals", () => {
  const en = translator("en-US");
  const de = translator("de-DE");
  assert.equal(en("summary.sources", { n: 1 }), "1 source");
  assert.equal(en("summary.sources", { n: 4 }), "4 sources");
  assert.equal(en("summary.sources", { n: 0 }), "0 sources");
  assert.equal(de("summary.citations", { n: 1 }), "1 Zitat");
  assert.equal(de("summary.citations", { n: 2 }), "2 Zitate");
  assert.equal(en("lookup.added", { key: "lee2010" }), "Added: lee2010");
  assert.equal(translator("fr")("refresh"), "Refresh");
  assert.equal(en("no.such.key"), "no.such.key");
});

test("document texts follow the document's locale", () => {
  assert.deepEqual(documentText("en-US"), { bibEmpty: "No sources cited.", citePlaceholder: "[empty citation]", bibTitle: "Bibliography" });
  assert.equal(documentText("de-DE").bibTitle, "Literaturverzeichnis");
  assert.deepEqual(documentText("ja-JP"), documentText("en-US"));
});

test("a UserError carries its key and reads in English", () => {
  const e = new UserError("error.doiNotFound", { reason: 404 });
  assert.equal(e.message, "DOI not found (404)");
  assert.equal(translator("de")(e.key, e.params), "DOI nicht gefunden (404)");
  assert.ok(e instanceof Error);
});
