// Runs one session of the binding experiment in a Word instance of its own
// (e2e/rig.mjs, e2e/word.ps1), driven by e2e/experiments/bind.js in the pane.
//
//   node e2e/experiments/bind.mjs <session> [--runs 3]
//
// Sessions (see SESSIONS): features, plain, plain-reopen, rich, fields,
// unbound, plain-20k, rich-20k, plain-pos. Each writes
// e2e/experiments/readings/<session>.json. A session whose input is another
// session's saved document reads it from dist-e2e/bind-<that>.docx.
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { build as esbuild } from "esbuild";
import { root, dist, build, requireWindows, writeDocx, readDocx, addPane, idFor } from "../rig.mjs";
import { inWord, stop } from "./word-exp.mjs";
import { generateBound, generateFeatures, generateMini, NS, PAGE } from "./generate-bound.mjs";
import { storeXml } from "./shared.js";

const LARGE = { pages: 500, citations: 4000, bibs: 6, refs: 2000 };
const SESSIONS = {
  // A first look at how the cost grows: one, 10 and 100 values, every writer, once.
  probe: { make: () => generateBound({ variant: "plain", ...LARGE }), steps: ["open", "timing"], runs: 1, counts: ["one", 10, 100, 400], writers: ["nodeText", "commonApi", "setXml"], all: ["nodeText", "commonApi"] },
  "probe-single": { make: () => generateBound({ variant: "plain", ...LARGE }), steps: ["open", "timing"], runs: 1, counts: ["one", 10, 100, 400], writers: ["nodeSingle"], all: ["nodeSingle"] },
  "probe-pos": { make: () => generateBound({ variant: "plain", xpath: "pos", ...LARGE }), steps: ["open", "timing"], runs: 1, counts: ["one", 10, 100, 400], writers: ["nodeSingle", "nodeText", "commonApi", "setXml"], all: ["nodeSingle"] },
  "probe-pos-medium": { make: () => generateBound({ variant: "plain", xpath: "pos", pages: 150, citations: 1000, bibs: 3, refs: 500 }), steps: ["open", "timing"], runs: 1, counts: ["one", 10, 100], writers: ["setXml", "updateElement", "nodeText", "commonApi"], all: ["setXml"] },
  "probe-medium": { make: () => generateBound({ variant: "plain", pages: 150, citations: 1000, bibs: 3, refs: 500 }), steps: ["open", "timing"], runs: 1, counts: ["one", 10, 100], writers: ["setXml", "updateElement"], all: ["setXml"] },
  mini: { make: () => generateMini(), steps: ["mini"] },
  features: { make: () => generateFeatures(), steps: ["features"] },
  plain: { make: () => generateBound({ variant: "plain", ...LARGE }), steps: ["open", "timing", "undo", "tracked"], counts: ["one", "tenth"], writers: ["nodeSingle", "commonApi"], all: ["nodeSingle"], allEveryRun: true },
  limits: { make: () => generateMini(), steps: ["limits"] },
  // The saved "plain" document with three values changed in the part only, as another tool would write it.
  "plain-reopen": { from: "plain", steps: ["open", "reopen"], edit: "store" },
  // The same with the texts changed in document.xml only.
  "plain-reopen-doc": { from: "plain", steps: ["open", "reopen"], edit: "document" },
  "plain-pos": { make: () => generateBound({ variant: "plain", xpath: "pos", ...LARGE }), steps: ["open", "timing"], runs: 1 },
  rich: { make: () => generateBound({ variant: "rich", ...LARGE }), steps: ["open", "timing"], counts: ["one", "tenth"], writers: ["nodeSingle"], all: ["nodeSingle"], allEveryRun: true },
  fields: { make: () => generateBound({ variant: "fields", ...LARGE }), steps: ["open", "fields"], runs: 1 },
  "fields-medium": { make: () => generateBound({ variant: "fields", pages: 150, citations: 1000, bibs: 3, refs: 500 }), steps: ["open", "fields"], runs: 1, quick: true },
  "rich-medium": { make: () => generateBound({ variant: "rich", pages: 150, citations: 1000, bibs: 3, refs: 500 }), steps: ["open", "timing"], counts: ["one", "tenth"], writers: ["nodeSingle"], all: ["nodeSingle"], allEveryRun: true },
  unbound: { make: () => generateBound({ variant: "unbound", ...LARGE }), steps: ["open", "paths"], runs: 3 },
  "plain-20k": { make: () => generateBound({ variant: "plain", ...LARGE, citations: 20000 }), steps: ["open", "timing"], runs: 1, counts: ["one", "tenth"], writers: ["nodeSingle", "setXml"], all: ["nodeSingle"] },
  "rich-20k": { make: () => generateBound({ variant: "rich", ...LARGE, citations: 20000 }), steps: ["open", "timing"], runs: 1 },
  "plain-unbound": { make: () => generateBound({ variant: "plain-unbound", ...LARGE }), steps: ["open", "direct"], runs: 3 },
  "body-ooxml": { make: () => generateBound({ variant: "unbound", ...LARGE }), steps: ["open", "bodyooxml"], runs: 3 },
  "unbound-open": { make: () => generateBound({ variant: "unbound", ...LARGE }), steps: ["open"] },
};

const argv = process.argv.slice(2);
const name = argv[0];
const session = SESSIONS[name];
if (!session) { console.error("sessions: " + Object.keys(SESSIONS).join(", ")); process.exit(2); }
const runsArg = argv.indexOf("--runs");
const runs = runsArg >= 0 ? +argv[runsArg + 1] : session.runs ?? 3;
const readings = path.join(root, "e2e/experiments/readings");
fs.mkdirSync(readings, { recursive: true });

requireWindows();
// The rig's build wipes `dist`, so a saved input document is kept aside first.
const saved = (s) => path.join(root, `e2e/experiments/.docs/bind-${s}.docx`);
build();
await esbuild({ entryPoints: { bind: path.join(root, "e2e/experiments/bind.js") }, outdir: dist, bundle: true, format: "iife", target: "es2020", sourcemap: "inline", logLevel: "warning" });
fs.writeFileSync(path.join(dist, PAGE), fs.readFileSync(path.join(root, "src/taskpane.html"), "utf8").replace('<script src="taskpane.js"></script>', '<script src="bind.js"></script>\n  <script src="taskpane.js"></script>'));

let parts, cites, stats;
if (session.from) {
  parts = readDocx(saved(session.from));
  const meta = JSON.parse(fs.readFileSync(path.join(readings, `${session.from}.json`), "utf8"));
  cites = meta.final;
  stats = { from: session.from, edit: session.edit };
  // Three citations, at the start, the middle and the end, changed in one place only.
  const pick = [0, Math.floor(cites.length / 2), cites.length - 1].map((i) => cites[i][0]);
  const item = Object.keys(parts).find((k) => /^customXml\/item\d+\.xml$/.test(k) && parts[k].includes(NS));
  if (session.edit === "store") {
    const values = new Map(cites);
    for (const id of pick) values.set(id, `[edited ${id}]`);
    parts[item] = storeXml(values);
    stats.expectedFromStore = pick;
  } else {
    let doc = parts["word/document.xml"];
    for (const id of pick) doc = doc.replace(new RegExp(`(<w:tag w:val="exp:cite:${id}"/>[\\s\\S]*?<w:sdtContent>[\\s\\S]*?<w:t[^>]*>)[^<]*`), `$1[doc-edited ${id}]`);
    parts["word/document.xml"] = doc;
    stats.expectedFromDocument = pick;
  }
} else {
  const g = session.make();
  parts = g.parts;
  cites = g.cites;
  stats = g.stats;
}
const docx = path.join(dist, `bind-${name}-in.docx`);
writeDocx(addPane(parts, idFor(PAGE)), docx);
stats.fileBytes = fs.statSync(docx).size;
console.log("generated " + JSON.stringify(stats));
fs.writeFileSync(path.join(dist, "bind-config.json"), JSON.stringify({ session: name, steps: session.steps, runs, cites, stats, counts: session.counts, writers: session.writers, all: session.all, allEveryRun: session.allEveryRun, quick: session.quick }));

const rows = [];
const t0 = Date.now();
// Every reading also goes to a .jsonl file as it arrives, so a stopped session keeps what it measured.
const partial = path.join(readings, `${name}.partial.jsonl`);
fs.writeFileSync(partial, JSON.stringify({ generated: stats }) + "\n");
const out = path.join(dist, `bind-${name}.docx`);
const r = await inWord({
  docx, out, page: PAGE, timeout: 4 * 3600,
  onMessage: (m) => {
    if (m.done) return;
    rows.push({ at: Date.now() - t0, ...m });
    fs.appendFileSync(partial, JSON.stringify(rows.at(-1)) + "\n");
    console.log(JSON.stringify(m).slice(0, 1500));
  },
}).finally(stop);
console.log("word " + JSON.stringify(r.info));

// What the saved file holds: the control texts in document.xml next to the part's values.
let file = null;
if (r.saved) {
  const p = readDocx(out);
  const doc = p["word/document.xml"];
  const item = Object.keys(p).find((k) => /^customXml\/item\d+\.xml$/.test(k) && p[k].includes(NS));
  const store = item ? new Map([...p[item].matchAll(/<c:c id="([^"]*)">([^<]*)<\/c:c>/g)].map((m) => [m[1], m[2]])) : null;
  const texts = new Map([...doc.matchAll(/<w:tag w:val="exp:cite:([^"]*)"\/>[\s\S]*?<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/g)]
    .map((m) => [m[1], [...m[2].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("")]));
  const ids = [...texts.keys()];
  const differ = store ? ids.filter((id) => store.has(id) && store.get(id).replace(/&amp;/g, "&") !== texts.get(id).replace(/&amp;/g, "&")) : [];
  const sample = [0, Math.floor(ids.length / 2), ids.length - 1].map((i) => ids[i]).filter(Boolean);
  file = {
    bytes: fs.statSync(out).size,
    documentXmlBytes: Buffer.byteLength(doc),
    storeParts: Object.keys(p).filter((k) => /^customXml\/item\d+\.xml$/.test(k)).length,
    controls: ids.length,
    textDiffersFromStore: differ.length,
    differExamples: differ.slice(0, 5).map((id) => [id, texts.get(id), store.get(id)?.slice(0, 80)]),
    sample: sample.map((id) => [id, texts.get(id), store?.get(id)?.slice(0, 80)]),
    dataBindings: (doc.match(/<w:dataBinding /g) ?? []).length,
    revisions: { ins: (doc.match(/<w:ins /g) ?? []).length, del: (doc.match(/<w:del /g) ?? []).length },
    authors: [...new Set([...doc.matchAll(/w:author="([^"]*)"/g)].map((m) => m[1]))].map((a) => (/^Author|^$/.test(a) ? a : "(another name)")),
    core: p["docProps/core.xml"] ? p["docProps/core.xml"].replace(/<dc:creator>[^<]*<\/dc:creator>/, (m) => (/Author|<dc:creator><\/dc:creator>/.test(m) ? m : "<dc:creator>(a name)</dc:creator>")) : null,
    appCompany: p["docProps/app.xml"] ? /<Company>/.test(p["docProps/app.xml"]) : null,
  };
  // The final values, for a session that reopens this document.
  if (store) file.final = ids.map((id) => [id, (store.get(id) ?? "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&")]);
  fs.mkdirSync(path.dirname(saved(name)), { recursive: true });
  fs.copyFileSync(out, saved(name));
}
const result = { session: name, steps: session.steps, runs, generated: stats, word: r.info, code: r.code, ok: r.done?.ok, seconds: (Date.now() - t0) / 1000, rows, file: file && { ...file, final: undefined }, final: name === "plain" ? file?.final : undefined };
fs.writeFileSync(path.join(readings, `${name}.json`), JSON.stringify(result, null, 1));
fs.rmSync(partial, { force: true });
fs.rmSync(docx, { force: true });
fs.rmSync(out, { force: true });
console.log(`done ${name}: exit ${r.code}, ok ${r.done?.ok}, ${result.seconds} s; file ${JSON.stringify(result.file)?.slice(0, 1500)}`);
