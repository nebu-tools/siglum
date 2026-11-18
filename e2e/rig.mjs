// What the Word harnesses (run.mjs, measure.mjs, probe.mjs) share: building
// the pages, making docx files that open a pane by themselves, and running
// one document in a Word instance of its own with the add-in sideloaded.
//
// A harness page talks to the harness with POST /log. A message
// {command: "<PowerShell>"} runs that script in word.ps1 against the open
// document and answers with what the script returned.
import { execFileSync, spawn } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { serve } from "../scripts/serve.mjs";

export const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
export const PS = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe";
// E2E_OUT and E2E_PORT let two harnesses run side by side.
export const PORT = +(process.env.E2E_PORT ?? 3187);
export const BASE = `http://localhost:${PORT}`;
export const dist = path.resolve(root, process.env.E2E_OUT ?? "dist-e2e");
export const ADDIN_ID = /<Id>([^<]+)<\/Id>/.exec(fs.readFileSync(path.join(root, "manifest.xml"), "utf8"))[1];

/** Exits with a message unless this is WSL on a Windows computer. */
export function requireWindows() {
  if (!fs.existsSync(PS)) {
    console.error("The Word harnesses need WSL on a Windows computer with Word.");
    process.exit(2);
  }
}

/** Builds the pane and every harness page (e2e.html, measure.html, probe.html) into `dist`. */
export function build() {
  execFileSync("node", [path.join(root, "scripts/build.mjs"), "--e2e", "--base", BASE, "--out", dist], { stdio: "inherit" });
}

/**
 * A manifest for one harness page. Each page gets an add-in id of its own, so
 * Word's add-in cache never serves one harness's page to another.
 */
export const idFor = (page) => (page === "e2e.html" ? ADDIN_ID : ADDIN_ID.slice(0, -12) + crypto.createHash("sha1").update(page).digest("hex").slice(0, 12));
function manifestFor(page) {
  const id = idFor(page);
  const xml = fs.readFileSync(path.join(dist, "manifest.xml"), "utf8")
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${id}</Id>`)
    .replaceAll(`${BASE}/e2e.html`, `${BASE}/${page}`);
  return { id, xml };
}

export const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export const EMPTY = {
  "[Content_Types].xml": `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  "_rels/.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": `${XML}<w:document ${W}><w:body><w:p/></w:body></w:document>`,
};

/** A document of plain paragraphs. */
export const textDocument = (...paragraphs) => ({
  ...EMPTY,
  "word/document.xml": `${XML}<w:document ${W}><w:body>${paragraphs.map((t) => `<w:p><w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`).join("") || "<w:p/>"}</w:body></w:document>`,
});

export const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Zips a map of part name → content into `file`. */
export function writeDocx(parts, file) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "siglum-docx-"));
  try {
    for (const [p, s] of Object.entries(parts)) {
      fs.mkdirSync(path.dirname(path.join(dir, p)), { recursive: true });
      fs.writeFileSync(path.join(dir, p), s);
    }
    fs.rmSync(file, { force: true });
    execFileSync("zip", ["-X", "-q", "-r", file, "."], { cwd: dir });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

/** The parts of a docx: XML parts as strings, everything else as Buffers. */
export function readDocx(file) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "siglum-docx-"));
  try {
    execFileSync("unzip", ["-q", file, "-d", dir]);
    const parts = {};
    const walk = (d) => {
      for (const e of fs.readdirSync(d, { withFileTypes: true })) {
        const p = path.join(d, e.name);
        if (e.isDirectory()) walk(p);
        else parts[path.relative(dir, p).split(path.sep).join("/")] = /\.(xml|rels)$/.test(e.name) ? fs.readFileSync(p, "utf8") : fs.readFileSync(p);
      }
    };
    walk(dir);
    return parts;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

/** Adds a task pane for add-in `id` that opens with the document, replacing any pane it had. */
export function addPane(parts, id = ADDIN_ID) {
  const out = Object.fromEntries(Object.entries(parts).filter(([k]) => !k.startsWith("word/webextensions/")));
  out["[Content_Types].xml"] = out["[Content_Types].xml"]
    .replace(/<Override PartName="\/word\/webextensions\/[^>]*\/>/g, "")
    .replace("</Types>", '<Override PartName="/word/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/><Override PartName="/word/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>');
  out["_rels/.rels"] = out["_rels/.rels"]
    .replace(/<Relationship [^>]*webextensiontaskpanes[^>]*\/>/g, "")
    .replace("</Relationships>", '<Relationship Id="rIdSiglumPane" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="word/webextensions/taskpanes.xml"/></Relationships>');
  out["word/webextensions/taskpanes.xml"] = `${XML}<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="right" visibility="1" width="350" row="0"><wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>`;
  out["word/webextensions/_rels/taskpanes.xml.rels"] = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/></Relationships>`;
  out["word/webextensions/webextension1.xml"] = `${XML}<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{0A1B2C3D-0000-4000-8000-000000000001}"><we:reference id="${id}" version="1.0.0.0" store="developer" storeType="Registry"/><we:alternateReferences/><we:properties><we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/></we:properties><we:bindings/><we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></we:webextension>`;
  return out;
}

/** Copies a docx, or writes `parts`, with a pane for the harness `page` that opens with it. */
export function withPane(source, file, page = "e2e.html") {
  const parts = typeof source === "string" ? readDocx(source) : source ?? EMPTY;
  writeDocx(addPane(parts, idFor(page)), file);
}

let winTemp = null;
const windowsTemp = () => (winTemp ??= execFileSync(PS, ["-NoProfile", "-Command", "$env:TEMP"], { encoding: "utf8" }).trim());

let server = null;
let route = () => null;
/** Starts the pane server once; `stop()` closes it. */
export async function start() {
  server ??= await serve(dist, PORT, (msg) => route(msg));
}
export function stop() {
  server?.close();
  server = null;
}

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Opens `docx` in a new Word instance with the harness `page` sideloaded and
 * calls `onMessage` for everything the page posts. The session ends when the
 * page posts {done: true}: Word then saves to `out` and quits. With
 * {done: true, save: false} it quits without saving.
 * Returns {code, saved, done, info}; `info` is what word.ps1 measured.
 */
export async function inWord({ docx, out, page = "e2e.html", onMessage = () => {}, timeout = 300, pages = false }) {
  await start();
  const { id, xml } = manifestFor(page);
  const name = `siglum-e2e-${process.pid}-${crypto.randomBytes(3).toString("hex")}`;
  const winDir = `${windowsTemp()}\\${name}`;
  const dir = execFileSync("wslpath", ["-u", winDir], { encoding: "utf8" }).trim();
  fs.mkdirSync(dir, { recursive: true });
  try {
    fs.writeFileSync(path.join(dir, "manifest.xml"), xml);
    fs.copyFileSync(path.join(root, "e2e/word.ps1"), path.join(dir, "word.ps1"));
    fs.copyFileSync(docx, path.join(dir, "in.docx"));
    let done = null;
    let commands = 0;
    let exited = false;
    route = async (msg) => {
      if (msg.command) {
        const n = String(++commands).padStart(5, "0");
        // PowerShell reads a script with a byte order mark as UTF-8.
        fs.writeFileSync(path.join(dir, `cmd-${n}.ps1`), "﻿" + msg.command);
        const ack = path.join(dir, `ack-${n}.json`);
        while (!fs.existsSync(ack)) {
          if (exited) return { error: "Word has quit" };
          await wait(100);
        }
        const answer = JSON.parse(fs.readFileSync(ack, "utf8").replace(/^﻿/, ""));
        fs.rmSync(ack);
        return answer;
      }
      if (msg.done) {
        done = msg;
        fs.writeFileSync(path.join(dir, msg.save === false ? "abort" : "done"), "");
      }
      return onMessage(msg);
    };
    const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", `${winDir}\\word.ps1`, "-Dir", winDir, "-AddinId", id, "-Timeout", String(timeout)];
    if (pages) args.push("-Pages");
    const code = await new Promise((resolve) => spawn(PS, args, { stdio: "inherit" }).on("exit", resolve));
    exited = true;
    route = () => null;
    const saved = fs.existsSync(path.join(dir, "out.docx"));
    if (saved && out) fs.copyFileSync(path.join(dir, "out.docx"), out);
    const infoFile = path.join(dir, "info.json");
    const info = fs.existsSync(infoFile) ? JSON.parse(fs.readFileSync(infoFile, "utf8").replace(/^﻿/, "")) : null;
    return { code, saved, done, info };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
