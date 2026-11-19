// Bundles the task pane into dist/ and writes a manifest for one web address.
//
//   node scripts/build.mjs [--base https://siglum.example.org] [--out dist] [--e2e]
import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const arg = (name, def) => {
  const i = process.argv.indexOf("--" + name);
  return i > 0 ? process.argv[i + 1] : def;
};
const e2e = process.argv.includes("--e2e");
const base = arg("base", process.env.SIGLUM_BASE ?? "https://localhost:3000").replace(/\/$/, "");
const out = path.resolve(root, arg("out", e2e ? "dist-e2e" : "dist"));
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
const entries = { taskpane: path.join(root, "src/taskpane.js") };
// Harness pages for Word (e2e/): each runs next to the real pane in a page of its own.
// scratch.js is for one-off experiments and is not committed.
const HARNESS = { e2e: "e2e/page.js", measure: "e2e/measure.js", probe: "e2e/probe.js", scratch: "e2e/scratch.js" };
if (e2e) for (const [name, file] of Object.entries(HARNESS)) if (fs.existsSync(path.join(root, file))) entries[name] = path.join(root, file);
await build({
  entryPoints: entries,
  outdir: out,
  bundle: true,
  format: "iife",
  target: "es2020",
  minify: !e2e,
  sourcemap: e2e ? "inline" : false,
  loader: { ".csl": "text", ".xml": "text" },
  logLevel: "warning",
});
for (const f of ["taskpane.html", "taskpane.css"]) fs.copyFileSync(path.join(root, "src", f), path.join(out, f));
for (const s of [16, 32, 64, 80]) fs.copyFileSync(path.join(root, `assets/icon-${s}.png`), path.join(out, `icon-${s}.png`));
if (e2e) {
  const html = fs.readFileSync(path.join(root, "src/taskpane.html"), "utf8");
  for (const name of Object.keys(HARNESS)) if (entries[name]) fs.writeFileSync(path.join(out, `${name}.html`), html.replace('<script src="taskpane.js"></script>', `<script src="${name}.js"></script>\n  <script src="taskpane.js"></script>`));
}
const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8")
  .replace(/^<!--.*-->\n/m, "")
  .replaceAll("{{BASE}}", base)
  .replaceAll("{{PAGE}}", e2e ? "e2e.html" : "taskpane.html")
  .replaceAll("{{VERSION}}", pkg.version + ".0");
fs.writeFileSync(path.join(out, "manifest.xml"), manifest);
console.log(`${path.relative(root, out)}/ for ${base}`);
