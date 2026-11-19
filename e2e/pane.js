// What the harness pages share (page.js, measure.js, probe.js). They run in
// the task pane next to the real pane, use it through its interface and talk
// to the harness through POST /log.
/* global Word, Office */

import { BIB_TAG, CITE_PREFIX, NS, storeFromXml } from "../src/model.js";

// The pages read the pane in English, whatever language Office has.
try { localStorage.setItem("siglum.language", "en"); } catch { /* Checks on status texts then fail and say so. */ }

export const post = (o) => fetch("/log", { method: "POST", body: JSON.stringify(o) }).then((r) => r.json()).catch(() => null);
export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
export const $ = (id) => document.getElementById(id);

/**
 * Runs a PowerShell script in word.ps1, where `$doc` is this document and
 * `$word` its Word instance, and returns what it returned. For what Office.js
 * cannot do, such as a tracked move or Word's Find settings.
 */
export async function command(script) {
  const r = await post({ command: script });
  if (!r || r.error) throw new Error("command failed: " + (r?.error ?? "no answer") + "\n" + script);
  return r.value;
}

// Every context.sync, from the pane or from a harness page, with its duration.
// `syncs.tag` says who is running, so a page can leave its own syncs out.
export const syncs = { log: [], tag: "pane" };
function patch() {
  const proto = globalThis.OfficeExtension?.ClientRequestContext?.prototype;
  if (!proto || proto.__counted) return;
  const sync = proto.sync;
  proto.sync = function (...a) {
    // Which function asked, from the stack: the frame below this one.
    const where = /at (?:async )?([\w$.]+) \(/.exec(String(new Error().stack).split("\n")[2] ?? "")?.[1];
    const e = { t: performance.now(), tag: syncs.tag, ms: 0, where };
    syncs.log.push(e);
    return sync.apply(this, a).finally(() => { e.ms = performance.now() - e.t; });
  };
  proto.__counted = true;
}
patch();
Office.onReady(() => patch());

/** Runs `fn` with Word as the harness, so its syncs do not count as the pane's. */
export async function word(fn) {
  const was = syncs.tag;
  syncs.tag = "harness";
  try { return await Word.run(fn); } finally { syncs.tag = was; }
}

/** The pane's syncs since `t`: count, and the time spent waiting for Word. */
export function paneSyncs(t, until = Infinity) {
  const s = syncs.log.filter((e) => e.tag === "pane" && e.t >= t && e.t <= until);
  return { syncs: s.length, syncMs: Math.round(s.reduce((n, e) => n + e.ms, 0)), each: s.map((e) => Math.round(e.ms)), where: s.map((e) => e.where ?? "?") };
}

// The pane announces every finished operation with "siglum:done".
const waiters = [];
export const done = { first: null, last: null };
document.addEventListener("siglum:done", (e) => {
  done.last = e.detail;
  done.first = done.first ?? e.detail;
  for (const w of waiters.splice(0)) w(e.detail);
});
/** The next operation the pane finishes. */
export const nextDone = (timeoutMs = 3 * 3600000) => new Promise((resolve, reject) => {
  waiters.push(resolve);
  setTimeout(() => reject(new Error(`the pane did not finish within ${timeoutMs} ms`)), timeoutMs);
});

/** Runs `fn` once the pane has finished its first refresh, then reports done. */
export function onReady(fn) {
  document.addEventListener("siglum:ready", async () => {
    let ok = true;
    try {
      ok = (await fn()) !== false;
    } catch (e) {
      ok = false;
      post({ error: String(e && (e.stack || e.message || e)) });
    }
    await sleep(200);
    post({ done: true, ok });
  }, { once: true });
}

/** Waits until the pane has not been busy for 300 ms. */
export async function idle() {
  await sleep(150);
  for (let calm = 0; calm < 3;) {
    await sleep(100);
    calm = document.body.hasAttribute("aria-busy") ? 0 : calm + 1;
  }
}

export function key(target, k, opts = {}) {
  target.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...opts }));
}

/** Types `query` into the search box and presses Enter, as a person citing does. */
export function pick(query, opts) {
  $("q").value = query;
  $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", opts);
}

/** Clicks Refresh and waits until the pane is done. */
export async function refresh() {
  const d = nextDone();
  $("refresh").click();
  const r = await d;
  await idle();
  return r;
}

export const heap = () => (performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null);

export const isCite = (x) => x.tag?.startsWith(CITE_PREFIX);
export const isBib = (x) => x.tag === BIB_TAG;

/** Citations and bibliographies in the body, in document order: [{kind, tag, text}]. */
export const controls = () => word(async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag,items/text");
  await c.sync();
  return all.items.filter((x) => isCite(x) || isBib(x)).map((x) => ({ kind: isBib(x) ? "bib" : "cite", tag: x.tag, text: x.text }));
});
export const citeTexts = async () => (await controls()).filter((x) => x.kind === "cite").map((x) => x.text);
export const bibTexts = async () => (await controls()).filter((x) => x.kind === "bib").map((x) => x.text.split(/[\r\n]+/).filter(Boolean));

/** The document's store as saved in its custom XML part, and how many such parts there are. */
export const readStore = () => word(async (c) => {
  const p = c.document.customXmlParts.getByNamespace(NS);
  p.load("items");
  await c.sync();
  if (!p.items.length) return { parts: 0, store: null };
  const x = p.items[0].getXml();
  await c.sync();
  return { parts: p.items.length, store: storeFromXml(x.value), bytes: x.value.length };
});

/** The tracked changes in the body: [[type, text]]. */
export const trackedChanges = () => word(async (c) => {
  const t = c.document.body.getTrackedChanges();
  t.load("items/type,items/text");
  await c.sync();
  return t.items.map((x) => [x.type, x.text]);
});

export const setTracking = (mode) => word(async (c) => { c.document.changeTrackingMode = mode; await c.sync(); });

/** Puts the cursor right after the first (or `nth`) occurrence of `text`. */
export const cursorAfter = (text, nth = 0) => word(async (c) => {
  const r = c.document.body.search(text, { matchCase: true });
  r.load("items");
  await c.sync();
  if (!r.items[nth]) throw new Error("not found: " + text);
  r.items[nth].getRange("End").select("End");
  await c.sync();
});

window.addEventListener("error", (e) => post({ error: String(e.message) }));
window.addEventListener("unhandledrejection", (e) => post({ error: String(e.reason && (e.reason.stack || e.reason)) }));
