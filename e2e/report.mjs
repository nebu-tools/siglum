// Turns measure.mjs results into one Markdown table, sizes side by side.
//
//   node e2e/report.mjs [dist-e2e/measure-small.json dist-e2e/measure-medium.json ...]
import fs from "node:fs";
import path from "node:path";
import { dist } from "./rig.mjs";

const files = process.argv.slice(2).length
  ? process.argv.slice(2)
  : ["small", "medium", "large"].map((s) => path.join(dist, `measure-${s}.json`)).filter((f) => fs.existsSync(f));
const runs = files.map((f) => JSON.parse(fs.readFileSync(f, "utf8")));
if (!runs.length) { console.error("no measure-*.json found"); process.exit(1); }

const sec = (ms) => (ms == null ? "" : ms >= 10000 ? `${Math.round(ms / 1000)} s` : `${(ms / 1000).toFixed(1)} s`);
const ops = [...new Set(runs.flatMap((r) => r.rows.map((x) => x.measure)))].filter((m) => m !== "error");
const head = runs.map((r) => `${r.size} (${r.word?.pages ?? r.spec.pages} p., ${r.generated.citations} cit., ${r.generated.bibs} lists, ${r.generated.refs} works)`);
console.log(`| operation | ${head.join(" | ")} |`);
console.log(`| --- |${runs.map(() => " --- |").join("")}`);
for (const op of ops) {
  const cells = runs.map((r) => {
    const x = r.rows.find((y) => y.measure === op);
    if (!x) return "";
    if (op.startsWith("read:")) return `${sec(x.readMs)} (${x.controls} controls)`;
    if (x.changes) return JSON.stringify(x.changes);
    if (x.bytes) return `${Math.round(x.bytes / 1024)} KB in ${x.parts} part(s)`;
    const parts = [sec(x.ms)];
    if (x.syncs != null) parts.push(`${x.syncs} syncs`);
    if (x.syncMs != null && x.ms) parts.push(`${Math.round((100 * x.syncMs) / x.ms)} % in Word`);
    if (x.writes) parts.push(`writes ${x.writes.citations}c/${x.writes.bibs}b`);
    if (x.refreshed === false) parts.push("no refresh");
    if (x.trackedChanges != null) parts.push(`${x.trackedChanges} tracked`);
    return parts.filter(Boolean).join(", ");
  });
  console.log(`| ${op} | ${cells.join(" | ")} |`);
}
const w = (f) => runs.map((r) => f(r) ?? "").join(" | ");
console.log(`| Word: open document (COM) | ${w((r) => sec(r.word?.openSeconds * 1000))} |`);
console.log(`| Word: save | ${w((r) => sec(r.word?.saveSeconds * 1000))} |`);
console.log(`| Word: peak working set | ${w((r) => r.word && `${r.word.peakWorkingSetMB} MB`)} |`);
console.log(`| Word: UI stalls ≥ 1 s (count, longest) | ${w((r) => r.word?.stalls && `${r.word.stalls.length}, ${sec(Math.max(0, ...r.word.stalls.map((s) => +s[1] || 0)))}`)} |`);
console.log(`| pane JS heap, last | ${w((r) => `${r.rows.at(-1)?.heap} MB`)} |`);
console.log(`| citeproc in Node, whole document | ${w((r) => sec(r.generated.nodeRenderMs))} |`);
