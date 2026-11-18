// Helpers for harness pages that run next to the real pane (measure.js,
// probe.js): talking to the harness, waiting for the pane, counting round trips.
/* global Word, OfficeExtension, Office */

export const post = (o) => fetch("/log", { method: "POST", body: JSON.stringify(o) }).catch(() => {});
export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
export const $ = (id) => document.getElementById(id);
export const word = (fn) => Word.run(fn);

// Every context.sync, from the pane or from the harness, with its duration.
// `syncs.tag` names who is running, so the harness can leave its own out.
export const syncs = { log: [], tag: "pane" };
function patch() {
  const proto = globalThis.OfficeExtension?.ClientRequestContext?.prototype;
  if (!proto || proto.__counted) return !!proto;
  const sync = proto.sync;
  proto.sync = function (...a) {
    const e = { t: performance.now(), tag: syncs.tag, ms: 0 };
    syncs.log.push(e);
    return sync.apply(this, a).finally(() => { e.ms = performance.now() - e.t; });
  };
  proto.__counted = true;
  return true;
}
patch();
Office.onReady(() => patch());

/** Runs `fn` as the harness, so its syncs are not counted as the pane's. */
export async function harness(fn) {
  const was = syncs.tag;
  syncs.tag = "harness";
  try { return await fn(); } finally { syncs.tag = was; }
}

/** The pane's syncs since `t`: count, and the time spent waiting for Word. */
export function paneSyncs(t, until = Infinity) {
  const s = syncs.log.filter((e) => e.tag === "pane" && e.t >= t && e.t <= until);
  return { syncs: s.length, syncMs: Math.round(s.reduce((n, e) => n + e.ms, 0)), each: s.map((e) => Math.round(e.ms)) };
}

// The pane announces every finished operation with "zitate:done" (or the
// renamed event); listen to every event on document whose name ends in ":done".
const waiters = [];
export let lastDone = null;
export let firstDone = null;
for (const name of ["zitate:done", "siglum:done", "citations:done"]) {
  document.addEventListener(name, (e) => {
    lastDone = e.detail;
    firstDone = firstDone ?? e.detail;
    for (const w of waiters.splice(0)) w(e.detail);
  });
}
export const nextDone = (timeoutMs = 3 * 3600000) => new Promise((resolve, reject) => {
  waiters.push(resolve);
  setTimeout(() => reject(new Error("pane did not finish in " + timeoutMs + " ms")), timeoutMs);
});
export function onReady(fn) {
  let ran = false;
  for (const name of ["zitate:ready", "siglum:ready", "citations:ready"]) {
    document.addEventListener(name, () => { if (!ran) { ran = true; fn(); } });
  }
}

/** Waits until the pane is idle (not aria-busy) for 300 ms. */
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

/** Types `query` into the search box and presses Enter, like a user citing. */
export function pick(query, opts) {
  $("q").value = query;
  $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter", opts);
}

export const heap = () => (performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null);

window.addEventListener("error", (e) => post({ error: String(e.message) }));
window.addEventListener("unhandledrejection", (e) => post({ error: String(e.reason && (e.reason.stack || e.reason)) }));
