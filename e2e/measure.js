// Measuring page: runs next to the real pane in a generated document
// (e2e/generate.mjs) and times the pane's own operations through its UI.
// Every result goes to the harness as {measure: name, ms, syncs, syncMs, ...}:
//   ms       wall time from the trigger until the pane reports done
//   syncs    the pane's context.sync calls in that time (round trips to Word)
//   syncMs   time spent inside those syncs, waiting for Word
//   writes   what the refresh rewrote: citations, lists, removed controls
// The harness serves its settings as measure-config.json.
/* global Office */
import { post, sleep, $, word, harness, paneSyncs, nextDone, onReady, idle, pick, heap, firstDone } from "./pane.js";
import * as names from "./names.mjs";
import { signature } from "../src/word.js";

const writes = (d) => d && { citations: d.citations?.length, bibs: d.bibs?.length, removes: d.removes?.length, cited: d.cited, pending: d.pending, unknown: d.unknown?.length };

function report(name, data) {
  post({ measure: name, heap: heap(), ...data });
}

/** Triggers a pane operation and measures it until the pane reports done. */
async function measure(name, trigger, extra = {}) {
  await idle();
  const t = performance.now();
  const done = nextDone();
  trigger();
  const d = await done;
  const ms = Math.round(performance.now() - t);
  report(name, { ms, ...paneSyncs(t), writes: writes(d), status: $("status").textContent.slice(0, 200), ...extra });
  await idle();
  return d;
}

const citeControls = async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag");
  await c.sync();
  return all.items.filter((x) => names.isCiteTag(x.tag));
};

/** Cursor at the start of the first text paragraph (after the first heading). */
const cursorAtStart = () => harness(() => word(async (c) => {
  c.document.body.paragraphs.getFirst().getNext().getRange("Start").select("Start");
  await c.sync();
}));

/** Cursor at the end of the last text paragraph before the last reference list heading. */
const cursorNearEnd = (bibs) => harness(() => word(async (c) => {
  const r = c.document.body.search(`Referenzliste ${bibs}`, { matchCase: true });
  r.load("items");
  await c.sync();
  r.items[r.items.length - 1].paragraphs.getFirst().getPrevious().getRange("Content").select("End");
  await c.sync();
}));

async function trackedChanges() {
  try {
    return await harness(() => word(async (c) => {
      const t = c.document.body.getTrackedChanges();
      t.load("items/type");
      await c.sync();
      const by = {};
      for (const x of t.items) by[x.type] = (by[x.type] ?? 0) + 1;
      return by;
    }));
  } catch (e) {
    return String(e.message ?? e);
  }
}

async function scenario(cfg) {
  // --steps in measure.mjs picks groups; without it every group runs.
  const on = (g) => !cfg.steps || cfg.steps.includes(g);
  // Opening: from page start to the pane's first refresh being done.
  report("open pane", { ms: Math.round(performance.now()), ...paneSyncs(0), writes: writes(firstDone), office: Office.context.diagnostics && { version: Office.context.diagnostics.version, platform: Office.context.diagnostics.platform } });

  if (on("idle")) {
    await measure("refresh, nothing to change", () => $("refresh").click());
  }

  if (on("watcher")) {
    // The selection watcher: a signature on its own, then a cursor move.
    {
      const t = performance.now();
      const sig = await harness(() => word((c) => signature(c)));
      report("signature (watcher's check)", { ms: Math.round(performance.now() - t), syncs: 1, length: sig.length });
    }
    {
      await cursorNearEnd(cfg.bibs);
      const t = performance.now();
      let fired = false;
      const d = nextDone(3500).then(() => (fired = true)).catch(() => {});
      await sleep(3500);
      await d;
      report("cursor move, nothing changed (3.5 s window)", { ms: null, ...paneSyncs(t), refreshed: fired });
    }
    {
      // Delete one citation in the middle as a user would, then move the cursor.
      await harness(() => word(async (c) => {
        const cc = await citeControls(c);
        cc[Math.floor(cc.length / 2)].delete(false);
        await c.sync();
      }));
      const t = performance.now();
      const done = nextDone(20 * 60000);
      await cursorAtStart();
      let d = null;
      try { d = await done; } catch (e) { report("cursor move after a deletion", { error: String(e.message) }); }
      if (d) report("cursor move after a deletion (watcher refresh)", { ms: Math.round(performance.now() - t), ...paneSyncs(t), writes: writes(d) });
      await idle();
    }
  }

  // Works no citation uses yet, so a citation at the start shifts every number.
  const [k1, k2, k3, k4] = cfg.fresh;

  if (on("cite")) {
    await cursorAtStart();
    await measure("cite from search at the start (all numbers shift)", () => pick(k1));

    await cursorNearEnd(cfg.bibs);
    await measure("cite from search near the end", () => pick(k2));
  }

  if (on("marker")) {
    const m = names.marker([k3]);
    if (m) {
      await harness(() => word(async (c) => {
        c.document.body.paragraphs.getFirst().getNext().getRange("Start").insertText(m + " ", "Before");
        await c.sync();
      }));
      await measure("typed marker at the start + refresh", () => $("refresh").click(), { marker: m });
    } else {
      report("typed marker at the start + refresh", { skipped: "no marker syntax found" });
    }
  }

  if (on("tracked")) {
    // Track changes on while a citation renumbers everything.
    await harness(() => word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); }));
    await cursorAtStart();
    await measure("cite at the start with track changes on", () => pick(k4));
    report("tracked changes after that", { changes: await trackedChanges() });
    await harness(() => word(async (c) => { c.document.changeTrackingMode = "Off"; await c.sync(); }));
  }

  if (on("full")) {
    // Every citation and list out of date: the full rewrite.
    {
      const t = performance.now();
      await harness(() => word(async (c) => {
        const cc = await citeControls(c);
        for (const x of cc) x.insertText(names.CITE_NEW, "Replace");
        await c.sync();
      }));
      report("(setup: reset every citation text)", { ms: Math.round(performance.now() - t) });
      await measure("full refresh, every citation rewritten", () => $("refresh").click());
    }
    await measure("refresh again, nothing to change", () => $("refresh").click());
  }

  if (on("search")) {
    // Search box: typing a query.
    {
      const t = performance.now();
      $("q").value = "mueller survival 2015";
      $("q").dispatchEvent(new Event("input"));
      report("search keystroke", { ms: Math.round(performance.now() - t), hits: document.querySelectorAll("#results li").length });
    }
  }

  if (on("reads")) {
    // Where a refresh's time goes: the snapshot reads every control's text with
    // getReviewedText twice. Each way of reading, on its own, in one sync.
    for (const [label, how] of [
      ["read: load tag,title,id,placeholderText", () => null],
      ["read: + load text", (x) => x.load("text")],
      ["read: + getReviewedText Current", (x) => x.getReviewedText("Current")],
      ["read: + getReviewedText Current and Original", (x) => [x.getReviewedText("Current"), x.getReviewedText("Original")]],
    ]) {
      const r = await harness(() => word(async (c) => {
        const all = c.document.body.contentControls;
        all.load("items/tag,items/title,items/id,items/placeholderText");
        const t0 = performance.now();
        await c.sync();
        const t1 = performance.now();
        all.items.forEach(how);
        await c.sync();
        return { loadMs: Math.round(t1 - t0), readMs: Math.round(performance.now() - t1), controls: all.items.length };
      }));
      report(label, r);
    }
    {
      const t = performance.now();
      const n = await harness(() => word(async (c) => {
        const tc = c.document.body.getTrackedChanges();
        tc.load("items/type");
        await c.sync();
        return tc.items.length;
      }));
      report("read: count tracked changes", { ms: Math.round(performance.now() - t), trackedChanges: n });
    }
  }

  const xml = await harness(() => word(async (c) => {
    const p = c.document.customXmlParts.getByNamespace(names.NS);
    p.load("items");
    await c.sync();
    const x = p.items[0].getXml();
    await c.sync();
    return { parts: p.items.length, bytes: x.value.length };
  }));
  report("stored library", xml);
}

onReady(async () => {
  let cfg = {};
  try { cfg = await (await fetch("measure-config.json", { cache: "no-store" })).json(); } catch { /* defaults */ }
  let ok = true;
  try {
    await scenario(cfg);
  } catch (e) {
    ok = false;
    post({ measure: "error", error: String(e && (e.stack || e.message || e)) });
  }
  post({ done: true, ok });
});
