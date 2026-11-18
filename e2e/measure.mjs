// Measures the add-in in real Word on generated documents (e2e/generate.mjs),
// one Word instance per size, driven by e2e/measure.js in the pane.
//
//   node e2e/measure.mjs                 small, medium and large
//   node e2e/measure.mjs small medium    just those
//   node e2e/measure.mjs --pages 50 --citations 300 --bibs 2 --refs 200
//
// Writes dist-e2e/measure-<size>.json and prints a table per size.
import fs from "node:fs";
import path from "node:path";
import { dist, build, requireWindows, inWord, writeDocx, addPane, idFor, stop } from "./rig.mjs";
import { generate } from "./generate.mjs";

const SIZES = {
  small: { pages: 20, citations: 100, bibs: 1, refs: 50 },
  medium: { pages: 150, citations: 1000, bibs: 3, refs: 500 },
  large: { pages: 500, citations: 4000, bibs: 6, refs: 2000 },
};

const argv = process.argv.slice(2);
const opt = (n) => { const i = argv.indexOf("--" + n); return i >= 0 ? +argv[i + 1] : undefined; };
let runs = argv.filter((a) => SIZES[a]).map((a) => [a, SIZES[a]]);
if (opt("pages")) runs = [["custom", { pages: opt("pages"), citations: opt("citations") ?? 100, bibs: opt("bibs") ?? 1, refs: opt("refs") ?? 50 }]];
if (!runs.length) runs = Object.entries(SIZES);

requireWindows();
build();

const fmt = (v) => (v == null ? "" : String(v));
try {
  for (const [size, spec] of runs) {
    console.log(`\n== ${size}: ${spec.pages} pages, ${spec.citations} citations, ${spec.bibs} lists, ${spec.refs} works`);
    const g = generate(spec);
    const docx = path.join(dist, `measure-${size}-in.docx`);
    writeDocx(addPane(g.parts, idFor("measure.html")), docx);
    console.log("    generated: " + JSON.stringify(g.stats));
    fs.writeFileSync(path.join(dist, "measure-config.json"), JSON.stringify({ ...g.stats, size }));
    const rows = [];
    const t0 = Date.now();
    const r = await inWord({
      docx, out: path.join(dist, `measure-${size}.docx`), page: "measure.html", timeout: 3 * 3600, pages: true,
      onMessage: (m) => {
        if (m.measure) {
          rows.push(m);
          const w = m.writes ? ` writes ${m.writes.citations}c/${m.writes.bibs}b/${m.writes.removes}r` : "";
          console.log(`    ${m.measure}: ${m.ms ?? ""} ms, ${fmt(m.syncs)} syncs (${fmt(m.syncMs)} ms in Word${m.each ? ": " + m.each.join(" ") : ""})${w}, heap ${fmt(m.heap)} MB${m.error ? " ERROR " + m.error : ""}`);
          if (m.ms == null && !m.writes) console.log("      " + JSON.stringify(m));
          else if (m.measure.startsWith("read:")) console.log("      " + JSON.stringify(m));
        } else if (!m.done) console.log("    " + JSON.stringify(m));
      },
    });
    const result = { size, spec, generated: g.stats, word: r.info, code: r.code, ok: r.done?.ok, seconds: (Date.now() - t0) / 1000, rows };
    fs.writeFileSync(path.join(dist, `measure-${size}.json`), JSON.stringify(result, null, 2));
    console.log(`    Word: ${JSON.stringify(r.info)}`);
    if (r.code !== 0 || !r.done?.ok) console.log(`FAIL ${size}: exit ${r.code}, done ${JSON.stringify(r.done)}`);
  }
} finally {
  stop();
}
