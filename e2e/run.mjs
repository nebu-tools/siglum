// End-to-end test in real Word, from WSL on a Windows machine with Word.
// Builds the e2e pane, sideloads it for one Word instance at a time, opens
// documents that open the pane by themselves and lets e2e/page.js drive it:
//
//   A  an empty document: citing, typed markers, bibliographies, deleting, track changes, DOI
//   B  a template with headings and placeholders: one bibliography per section
//   C  Word reopens B's result: everything still current, citing goes on
//
//   npm run e2e        results in dist-e2e/{a,b,c}.docx
import { execFileSync, spawn } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { serve } from "../scripts/serve.mjs";
import { BIB_TAG, CITE_PREFIX } from "../src/model.js";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const PS = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe";
const PORT = 3187;
const ADDIN_ID = /<Id>([^<]+)<\/Id>/.exec(fs.readFileSync(path.join(root, "manifest.xml"), "utf8"))[1];

if (!fs.existsSync(PS)) {
  console.error("The e2e tests need WSL on a Windows machine with Word.");
  process.exit(2);
}

execFileSync("node", [path.join(root, "scripts/build.mjs"), "--e2e", "--base", `http://localhost:${PORT}`], { stdio: "inherit" });
const dist = path.join(root, "dist-e2e");

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const document = (text) => `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ""}</w:p></w:body></w:document>`;
const EMPTY = {
  "[Content_Types].xml": `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  "_rels/.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
};

/**
 * Copies a docx, or makes one with a single paragraph of `text`, and adds a
 * task pane that opens with it. The text tells e2e/page.js which scenario to run.
 */
function withPane({ base, text = "" }, file) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "siglum-docx-"));
  if (base) execFileSync("unzip", ["-q", base, "-d", dir]);
  const put = (p, s) => { fs.mkdirSync(path.dirname(path.join(dir, p)), { recursive: true }); fs.writeFileSync(path.join(dir, p), s); };
  if (!base) for (const [p, s] of Object.entries({ ...EMPTY, "word/document.xml": document(text) })) put(p, s);
  const ct = path.join(dir, "[Content_Types].xml");
  fs.writeFileSync(ct, fs.readFileSync(ct, "utf8").replace("</Types>", '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/><Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>'));
  const rels = path.join(dir, "_rels/.rels");
  fs.writeFileSync(rels, fs.readFileSync(rels, "utf8").replace("</Relationships>", '<Relationship Id="rIdSiglumPane" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="word/webextensions/taskpanes.xml"/></Relationships>'));
  put("word/webextensions/taskpanes.xml", `${XML}<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="right" visibility="1" width="350" row="0"><wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>`);
  put("word/webextensions/_rels/taskpanes.xml.rels", `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/></Relationships>`);
  put("word/webextensions/webextension1.xml", `${XML}<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{0A1B2C3D-0000-4000-8000-000000000001}"><we:reference id="${ADDIN_ID}" version="1.0.0.0" store="developer" storeType="Registry"/><we:alternateReferences/><we:properties><we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/></we:properties><we:bindings/><we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></we:webextension>`);
  fs.rmSync(file, { force: true });
  execFileSync("zip", ["-X", "-q", "-r", file, "."], { cwd: dir });
  fs.rmSync(dir, { recursive: true });
}

const win = execFileSync(PS, ["-NoProfile", "-Command", "$env:TEMP"], { encoding: "utf8" }).trim();
const winDir = `${win}\\siglum-e2e`;
const dir = execFileSync("wslpath", ["-u", winDir], { encoding: "utf8" }).trim();
let log = () => {};
const server = await serve(dist, PORT, (msg) => log(msg));

/** Opens `docx` in Word with the e2e pane and returns its checks and the saved file. */
async function inWord(name, docx, out) {
  console.log(`\n== ${name}`);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  fs.copyFileSync(path.join(dist, "manifest.xml"), path.join(dir, "manifest.xml"));
  fs.copyFileSync(path.join(root, "e2e/word.ps1"), path.join(dir, "word.ps1"));
  fs.copyFileSync(docx, path.join(dir, "in.docx"));
  const checks = [];
  let done = null;
  log = (msg) => {
    if (msg.check) {
      checks.push(msg);
      console.log(`${msg.ok ? "ok  " : "FAIL"} ${msg.check}${msg.ok ? "" : "  " + JSON.stringify(msg.detail)}`);
    } else if (msg.done) {
      done = msg;
      fs.writeFileSync(path.join(dir, "done"), "");
    } else {
      console.log("    " + JSON.stringify(msg));
    }
  };
  const code = await new Promise((resolve) => {
    const p = spawn(PS, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", `${winDir}\\word.ps1`, "-Dir", winDir, "-AddinId", ADDIN_ID], { stdio: "inherit" });
    p.on("exit", resolve);
  });
  const saved = fs.existsSync(path.join(dir, "out.docx"));
  if (saved) fs.copyFileSync(path.join(dir, "out.docx"), out);
  fs.rmSync(dir, { recursive: true, force: true });
  const ok = code === 0 && saved && done?.ok === true && checks.length > 0;
  if (!ok) console.log(`FAIL ${name}: ${done?.failed?.join(", ") || (saved ? "the page did not finish" : "the document was not saved")}`);
  return ok;
}

function check(name, pass, detail) {
  console.log(`${pass ? "ok  " : "FAIL"} ${name}${pass ? "" : "  " + JSON.stringify(detail)}`);
  return pass;
}

const unzipped = (docx, part) => execFileSync("unzip", ["-p", docx, part], { encoding: "utf8" });
const tagged = (xml, tag) => (xml.match(new RegExp(`<w:tag w:val="${tag}`, "g")) ?? []).length;
const file = (n) => path.join(dist, n);
let ok = true;
try {
  withPane({}, file("a-in.docx"));
  ok &&= await inWord("A: an empty document", file("a-in.docx"), file("a.docx"));
  if (ok) {
    const doc = unzipped(file("a.docx"), "word/document.xml");
    ok &&= check("A saved: citations as content controls", tagged(doc, CITE_PREFIX) >= 8, tagged(doc, CITE_PREFIX));
    ok &&= check("A saved: two bibliographies", tagged(doc, BIB_TAG + '"') === 2, tagged(doc, BIB_TAG + '"'));
  }

  withPane({ text: "e2e:template" }, file("b-in.docx"));
  ok &&= await inWord("B: a template", file("b-in.docx"), file("b.docx"));
  if (ok) ok &&= await inWord("C: B's result, opened again", file("b.docx"), file("c.docx"));

} finally {
  server.close();
}
console.log(ok ? "\nE2E passed." : "\nE2E failed.");
process.exit(ok ? 0 : 1);
