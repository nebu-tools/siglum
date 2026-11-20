// The end-to-end suite in real Word (npm run e2e). Each scenario opens a
// document in a Word instance of its own; e2e/page.js in the pane drives the
// real pane, checks what Word holds and reports pass or fail per check. The
// harness then checks the saved file.
//
//   npm run e2e                 every scenario
//   npm run e2e -- markers      only those, by name
//
// The results are in dist-e2e/<scenario>.docx.
import fs from "node:fs";
import path from "node:path";
import { dist, build, requireWindows, withPane, inWord, readDocx, stop } from "./rig.mjs";
import { BIB_TAG, CITE_PREFIX, NS, storeFromXml } from "../src/model.js";
import { generate } from "./generate.mjs";

requireWindows();
build();

const file = (n) => path.join(dist, n);
const results = [];

function check(name, pass, detail, known) {
  results.push({ name, ok: !!pass || !!known, known: known && !pass });
  const mark = pass ? "ok  " : known ? "KNOWN" : "FAIL";
  console.log(`${mark} ${name}${pass ? "" : "  " + JSON.stringify(detail)}${known && !pass ? "\n      " + known : ""}`);
  return !!pass || !!known;
}

/** Opens `docx` with the e2e page, which runs `scenario`. Returns whether all its checks passed. */
async function session(name, scenario, docx, out) {
  console.log(`\n== ${name}`);
  fs.writeFileSync(file("e2e-config.json"), JSON.stringify({ scenario }));
  let checks = 0;
  let failed = 0;
  const r = await inWord({
    docx, out, page: "e2e.html", timeout: 900,
    onMessage: (msg) => {
      if (msg.check) {
        checks++;
        if (!check(`${scenario}: ${msg.check}`, msg.ok, msg.detail, msg.known)) failed++;
      } else if (!msg.done) {
        console.log("    " + JSON.stringify(msg));
      }
    },
  });
  if (r.info) console.log(`    Word ${r.info.version} build ${r.info.build}, pid ${r.info.pid}, open ${r.info.openSeconds?.toFixed(1)} s, pane ${r.info.paneSeconds?.toFixed(1)} s`);
  const ok = r.code === 0 && (r.saved || !out) && r.done?.ok === true && checks > 0 && !failed;
  if (!ok) check(`${scenario}: session`, false, { code: r.code, saved: r.saved, done: r.done, checks });
  return ok;
}

const saved = (n) => readDocx(file(n));
const tagged = (xml, tag) => (xml.match(new RegExp(`<w:tag w:val="${tag}`, "g")) ?? []).length;
const storeOf = (parts) => {
  const xml = Object.entries(parts).filter(([k, v]) => /^customXml\/item\d+\.xml$/.test(k) && typeof v === "string" && v.includes(NS));
  return { parts: xml.length, store: xml.length ? storeFromXml(xml[0][1]) : null };
};

const empty = (name) => () => withPane(null, file(`${name}-in.docx`));
const again = (from) => (name) => withPane(file(`${from}.docx`), file(`${name}-in.docx`));
// [name, what it does, how its input is made]. A later one opens an earlier one's result.
const SCENARIOS = [
  ["main", "an empty document: citing, markers, bibliographies, deleting, track changes, lookup", empty("main")],
  ["main-reopen", "main's result, saved, closed and opened again", again("main")],
  ["template", "a template with headings and placeholders", empty("template")],
  ["template-reopen", "template's result, opened again", again("template")],
  ["markers", "typed markers next to fields, tracked changes, comments, controls and in long groups", empty("markers")],
  ["sections", "copy, cut and move of citations between sections, tracked moves", empty("sections")],
  ["sections-reopen", "sections' result, opened again", again("sections")],
  ["renumber", "a document of 300 citations renumbered in batches: a second action, a deletion and a closed pane meanwhile", (name) => withPane(generate({ pages: 30, citations: 300, bibs: 2, refs: 120, seed: 3 }).parts, file(`${name}-in.docx`))],
  ["renumber-reopen", "renumber's result, saved while the batches ran, opened again", again("renumber")],
  ["fastpath", "a single-section document of 600 citations renumbered in one body write: a text box, a person's bookmark and a comment thread", (name) => withPane(generate({ pages: 60, citations: 600, bibs: 1, refs: 200, seed: 7 }).parts, file(`${name}-in.docx`))],
  ["edit", "citing, then editing the citation's locator in the pane and removing it", (name) => withPane(generate({ pages: 2, citations: 2, bibs: 1, refs: 10, seed: 5 }).parts, file(`${name}-in.docx`))],
];

async function scenario([name, what, make]) {
  if (make === undefined) return;
  try { make(name); } catch (e) { check(`${name}: input`, false, String(e.message)); return; }
  if (!await session(what, name, file(`${name}-in.docx`), file(`${name}.docx`))) return;
  const parts = saved(`${name}.docx`);
  const doc = parts["word/document.xml"];
  const { parts: n } = storeOf(parts);
  check(`${name}: the saved file holds one store`, n === 1, n);
  if (name === "main") {
    check("main: saved citations are content controls", tagged(doc, CITE_PREFIX) >= 8, tagged(doc, CITE_PREFIX));
    check("main: saved with two bibliographies", tagged(doc, BIB_TAG + '"') === 2, tagged(doc, BIB_TAG + '"'));
  }
  if (name === "fastpath") {
    check("fastpath: the person's bookmark is saved", doc.includes('w:name="keepMe"'), doc.match(/<w:bookmarkStart\b[^>]*>/)?.[0]);
    check("fastpath: the text box text is saved", doc.includes("shape keeps"), doc.includes("shape keeps"));
  }
}

const want = process.argv.slice(2);
for (const w of want) {
  if (!SCENARIOS.some(([n]) => n === w)) check(`${w}: input`, false, `no such scenario (one of ${SCENARIOS.map(([n]) => n).join(", ")})`);
}
// A reopen scenario needs its first run's document: include it.
for (const [n] of SCENARIOS) {
  const first = n.replace(/-reopen$/, "");
  if (first !== n && want.includes(n) && !want.includes(first)) want.unshift(first);
}
try {
  for (const sc of SCENARIOS) if (!want.length || want.includes(sc[0])) await scenario(sc);
} finally {
  stop();
}
const failed = results.filter((r) => !r.ok);
const knownFails = results.filter((r) => r.known);
const knownLine = knownFails.length ? `\n${knownFails.length} known defect(s), listed in the README's limits:\n${knownFails.map((r) => "  " + r.name).join("\n")}` : "";
console.log((failed.length ? `\nE2E failed: ${failed.length} of ${results.length} checks.\n${failed.map((r) => "  " + r.name).join("\n")}` : `\nE2E passed: ${results.length} checks.`) + knownLine);
process.exit(failed.length ? 1 : 0);
