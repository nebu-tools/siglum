// The experiment's inWord(): e2e/rig.mjs's, but with e2e/experiments/word-exp.ps1,
// which identifies its Word process before it opens the document.
import { execFileSync, spawn } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { serve } from "../../scripts/serve.mjs";
import { root, PS, PORT, dist, idFor } from "../rig.mjs";

let server = null;
let route = () => null;
export async function start() { server ??= await serve(dist, PORT, (msg) => route(msg)); }
export function stop() { server?.close(); server = null; }
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
let winTemp = null;
const windowsTemp = () => (winTemp ??= execFileSync(PS, ["-NoProfile", "-Command", "$env:TEMP"], { encoding: "utf8" }).trim());

function manifestFor(page) {
  const id = idFor(page);
  const xml = fs.readFileSync(path.join(dist, "manifest.xml"), "utf8").replace(/<Id>[^<]+<\/Id>/, `<Id>${id}</Id>`).replaceAll(/http:\/\/localhost:\d+\/e2e\.html/g, `http://localhost:${PORT}/${page}`);
  return { id, xml };
}

/** As rig.mjs's inWord. `docs` are extra files copied into the session folder. */
export async function inWord({ docx, out, page, onMessage = () => {}, timeout = 300, docs = {} }) {
  await start();
  const { id, xml } = manifestFor(page);
  const name = `siglum-e2e-${process.pid}-${crypto.randomBytes(3).toString("hex")}`;
  const winDir = `${windowsTemp()}\\${name}`;
  const dir = execFileSync("wslpath", ["-u", winDir], { encoding: "utf8" }).trim();
  fs.mkdirSync(dir, { recursive: true });
  try {
    fs.writeFileSync(path.join(dir, "manifest.xml"), xml);
    fs.copyFileSync(path.join(root, "e2e/experiments/word-exp.ps1"), path.join(dir, "word.ps1"));
    fs.copyFileSync(docx, path.join(dir, "in.docx"));
    for (const [n, f] of Object.entries(docs)) fs.copyFileSync(f, path.join(dir, n));
    let done = null;
    let commands = 0;
    let exited = false;
    route = async (msg) => {
      if (msg.command) {
        const n = String(++commands).padStart(5, "0");
        fs.writeFileSync(path.join(dir, `cmd-${n}.ps1`), "﻿" + msg.command);
        const ack = path.join(dir, `ack-${n}.json`);
        while (!fs.existsSync(ack)) {
          if (exited) return { error: "Word has quit" };
          await wait(50);
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
    const alive = path.join(dir, "alive");
    fs.writeFileSync(alive, "");
    const beat = setInterval(() => { try { const t = new Date(); fs.utimesSync(alive, t, t); } catch { /* over */ } }, 5000);
    const code = await new Promise((resolve) => spawn(PS, args, { stdio: "inherit", detached: true }).on("exit", resolve));
    clearInterval(beat);
    exited = true;
    route = () => null;
    const saved = fs.existsSync(path.join(dir, "out.docx"));
    if (saved && out) fs.copyFileSync(path.join(dir, "out.docx"), out);
    const infoFile = path.join(dir, "info.json");
    const info = fs.existsSync(infoFile) ? JSON.parse(fs.readFileSync(infoFile, "utf8").replace(/^﻿/, "")) : null;
    return { code, saved, done, info, dir: winDir };
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
