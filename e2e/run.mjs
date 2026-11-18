// End-to-end test in real Word, from WSL on a Windows machine with Word.
// Builds the e2e pane, sideloads it for one Word instance at a time, opens
// documents that open the pane by themselves and lets e2e/page.js drive it:
//
//   A  an empty document: citing, markers, lists, deleting, track changes, DOI
//   B  the G-BA Modul 3 template: lists in two "Referenzliste" sections
//   R  `gba.R apply` rewrites B's result, as the pipeline does with team files
//   C  Word reopens R's result: everything still current, citing goes on
//
//   npm run e2e        results in dist-e2e/{a,b,r,c}.docx
//
// Needs a tabellion checkout for gba.R and the templates: ../tabellion, or set TABELLION_DIR.
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { root, dist, build, requireWindows, withPane, inWord as session, stop } from "./rig.mjs";

const repo = path.resolve(process.env.TABELLION_DIR ?? path.join(root, "..", "tabellion"));
const templates = path.join(repo, "templates");
const TEMPLATE = fs.existsSync(templates) ? fs.readdirSync(templates).find((n) => /Modul3\.docx$/.test(n)) : undefined;

requireWindows();
if (!TEMPLATE) {
  console.error(`Modul-3-Vorlage fehlt in ${templates}: erst \`rv run gba.R fetch\` in tabellion, oder TABELLION_DIR setzen.`);
  process.exit(2);
}
build();

/** Opens `docx` in Word with the e2e pane and returns whether all its checks passed. */
async function inWord(name, docx, out) {
  console.log(`\n== ${name}`);
  const checks = [];
  const r = await session({
    docx, out, page: "e2e.html",
    onMessage: (msg) => {
      if (msg.check) {
        checks.push(msg);
        console.log(`${msg.ok ? "ok  " : "FAIL"} ${msg.check}${msg.ok ? "" : "  " + JSON.stringify(msg.detail)}`);
      } else if (!msg.done) {
        console.log("    " + JSON.stringify(msg));
      }
    },
  });
  if (r.info) console.log(`    Word ${r.info.version} build ${r.info.build}, open ${r.info.openSeconds?.toFixed(1)} s, pane ${r.info.paneSeconds?.toFixed(1)} s`);
  const ok = r.code === 0 && r.saved && r.done?.ok === true && checks.length > 0;
  if (!ok) console.log(`FAIL ${name}: ${r.done?.failed?.join(", ") || (r.saved ? "kein Abschluss" : "nicht gespeichert")}`);
  return ok;
}

function check(name, pass, detail) {
  console.log(`${pass ? "ok  " : "FAIL"} ${name}${pass ? "" : "  " + JSON.stringify(detail)}`);
  return pass;
}

const outline = (docx) => JSON.parse(execFileSync("rv", ["run", "gba.R", "outline", docx], { cwd: repo, encoding: "utf8" }));
const file = (n) => path.join(dist, n);
let ok = true;
try {
  withPane(null, file("a-in.docx"));
  ok &&= await inWord("A: leeres Dokument", file("a-in.docx"), file("a.docx"));
  if (ok) {
    const doc = execFileSync("unzip", ["-p", file("a.docx"), "word/document.xml"], { encoding: "utf8" });
    ok &&= check("A saved: citations as content controls", (doc.match(/<w:tag w:val="zitat:/g) ?? []).length >= 7);
    ok &&= check("A saved: two reference lists", (doc.match(/<w:tag w:val="literaturverzeichnis"/g) ?? []).length === 2);
  }

  withPane(path.join(repo, "templates", TEMPLATE), file("b-in.docx"));
  ok &&= await inWord("B: Modul-3-Vorlage", file("b-in.docx"), file("b.docx"));

  if (ok) {
    console.log("\n== R: gba.R apply (Felder) und outline");
    const ops = file("ops.json");
    fs.writeFileSync(ops, JSON.stringify({ aenderungen: [{ op: "felder" }] }));
    execFileSync("rv", ["run", "gba.R", "apply", file("b.docx"), ops, file("r.docx")], { cwd: repo, stdio: ["ignore", "ignore", "inherit"] });
    const o = outline(file("r.docx"));
    ok &&= check("R outline: citations", JSON.stringify(o.zitate.map((z) => [z.schluessel, z.text])) === JSON.stringify([[["mueller2020"], "[1]"], [["csr"], "[2]"], [["lee2010", "mueller2020"], "[1,2]"]]), o.zitate);
    ok &&= check("R outline: two lists, four sources", o.literaturverzeichnisse.length === 2 && o.quellen === 4, [o.literaturverzeichnisse, o.quellen]);
    ok &&= await inWord("C: R-Ergebnis wieder in Word", file("r.docx"), file("c.docx"));
  }
} finally {
  stop();
}
console.log(ok ? "\nE2E bestanden." : "\nE2E fehlgeschlagen.");
process.exit(ok ? 0 : 1);
