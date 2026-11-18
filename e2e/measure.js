// The measuring page: runs next to the real pane in a generated document
// (e2e/generate.mjs) and times the pane's own operations through its interface.
// Each result goes to e2e/measure.mjs as {measure: name, ms, syncs, syncMs, ...}:
//   ms        wall time from the trigger until the pane reports done
//   visibleMs for a citation from search, until the pane says the citation
//             at the cursor has its number ("siglum:visible"), if it says so
//   settledMs until the pane says the whole document is current ("siglum:settled")
//   syncs     the pane's context.sync calls in that time
//   syncMs    time spent inside those syncs, waiting for Word
//   writes    what the operation wrote: citations, bibliographies, removed controls
// The settings come from measure-config.json.
/* global Office */
import { post, sleep, $, word, command, paneSyncs, nextDone, done, onReady, idle, pick, heap, isCite, isBib, readStore, setTracking, trackedChanges } from "./pane.js";
import { CITE_NEW } from "../src/model.js";
import { signature } from "../src/word.js";

const writes = (d) => d && { citations: d.citations?.length, bibs: d.bibs?.length, removes: d.removes?.length, cited: d.cited, pending: d.pending, unknown: d.unknown?.length };
const report = (name, data) => post({ measure: name, heap: heap(), ...data });

let visibleAt = null;
let settledAt = null;
document.addEventListener("siglum:visible", () => { visibleAt ??= performance.now(); });
document.addEventListener("siglum:settled", () => { settledAt = performance.now(); });

/** Triggers a pane operation and measures it until the pane reports done. */
async function measure(name, trigger, extra = {}) {
  await idle();
  const t = performance.now();
  visibleAt = null;
  settledAt = null;
  const d = nextDone();
  trigger();
  const r = await d;
  // A pane error is not a completed operation: fail the session instead of
  // recording it as one.
  if (r?.error) throw new Error(`pane error during ${name}: ${String(r.error).slice(0, 500)}`);
  const ms = Math.round(performance.now() - t);
  // A pane that goes on writing in batches after "done" says "siglum:settled" when the document is current.
  await idle();
  const settled = Math.round((settledAt ?? performance.now()) - t);
  report(name, { ms, settledMs: settled, visibleMs: visibleAt == null ? undefined : Math.round(visibleAt - t), ...paneSyncs(t), writes: writes(r), status: $("status").textContent.slice(0, 200), ...extra });
  return r;
}

const citations = async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag");
  await c.sync();
  return all.items.filter(isCite);
};

/** The cursor at the start of the first text paragraph, after the first heading. */
const cursorAtStart = () => word(async (c) => {
  c.document.body.paragraphs.getFirst().getNext().getRange("Start").select("Start");
  await c.sync();
});

/** The cursor at the end of the paragraph before the last bibliography's heading. */
const cursorNearEnd = (bibs) => word(async (c) => {
  const r = c.document.body.search(`References ${bibs}`, { matchCase: true });
  r.load("items");
  await c.sync();
  r.items[r.items.length - 1].paragraphs.getFirst().getPrevious().getRange("Content").select("End");
  await c.sync();
});

async function scenario(cfg) {
  // --steps in measure.mjs picks groups; without it every group runs.
  const on = (g) => !cfg.steps || cfg.steps.includes(g);
  if (done.first?.error) throw new Error(`pane error on open: ${String(done.first.error).slice(0, 500)}`);
  report("open pane", { ms: Math.round(performance.now()), ...paneSyncs(0), writes: writes(done.first), office: Office.context.diagnostics && { version: Office.context.diagnostics.version, platform: Office.context.diagnostics.platform } });

  if (on("idle")) await measure("refresh, nothing to change", () => $("refresh").click());

  if (on("watcher")) {
    {
      const t = performance.now();
      const sig = await word((c) => signature(c));
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
      // Delete one citation in the middle as a person would, then move the cursor.
      await word(async (c) => {
        const cc = await citations(c);
        cc[Math.floor(cc.length / 2)].delete(false);
        await c.sync();
      });
      const t = performance.now();
      const d = nextDone(3 * 3600000);
      await cursorAtStart();
      const r = await d;
      if (r?.error) throw new Error(`pane error after a deletion: ${String(r.error).slice(0, 500)}`);
      report("cursor move after a deletion (watcher refresh)", { ms: Math.round(performance.now() - t), ...paneSyncs(t), writes: writes(r) });
      await idle();
    }
  }

  // Sources that no citation uses yet, so a citation at the start shifts every number.
  const [k1, k2, k3, k4] = cfg.fresh;

  if (on("cite")) {
    await cursorAtStart();
    await measure("cite from search at the start (all numbers shift)", () => pick(k1));
    await cursorNearEnd(cfg.bibs);
    await measure("cite from search near the end", () => pick(k2));
  }

  if (on("marker")) {
    await word(async (c) => {
      c.document.body.paragraphs.getFirst().getNext().getRange("Start").insertText(`@${k3} `, "Before");
      await c.sync();
    });
    await measure("typed marker at the start + refresh", () => $("refresh").click());
  }

  if (on("tracked")) {
    await setTracking("TrackAll");
    await cursorAtStart();
    await measure("cite at the start with track changes on", () => pick(k4));
    const tc = await trackedChanges().catch((e) => String(e.message ?? e));
    report("tracked changes after that", { changes: Array.isArray(tc) ? tc.reduce((by, [type]) => ({ ...by, [type]: (by[type] ?? 0) + 1 }), {}) : tc });
    await setTracking("Off");
    await measure("refresh with a tracked change, nothing to change", () => $("refresh").click());
  }

  if (on("full")) {
    const t = performance.now();
    await word(async (c) => {
      for (const x of await citations(c)) x.insertText(CITE_NEW, "Replace");
      await c.sync();
    });
    report("(setup: reset every citation text)", { ms: Math.round(performance.now() - t) });
    await measure("full refresh, every citation rewritten", () => $("refresh").click());
    await measure("refresh again, nothing to change", () => $("refresh").click());
  }

  if (on("search")) {
    const t = performance.now();
    $("q").value = "summer temperature 2015";
    $("q").dispatchEvent(new Event("input"));
    report("search keystroke", { ms: Math.round(performance.now() - t), hits: document.querySelectorAll("#results li").length });
  }

  if (on("reads")) {
    // Each way of reading every control's text, on its own, in one sync.
    for (const [label, how] of [
      ["read: load tag,title,id,placeholderText", () => null],
      ["read: + load text", (x) => x.load("text")],
      ["read: + getReviewedText Current", (x) => x.getReviewedText("Current")],
      ["read: + getReviewedText Current and Original", (x) => [x.getReviewedText("Current"), x.getReviewedText("Original")]],
    ]) {
      const r = await word(async (c) => {
        const all = c.document.body.contentControls;
        all.load("items/tag,items/title,items/id,items/placeholderText");
        const t0 = performance.now();
        await c.sync();
        const t1 = performance.now();
        all.items.forEach(how);
        await c.sync();
        return { loadMs: Math.round(t1 - t0), readMs: Math.round(performance.now() - t1), controls: all.items.length };
      });
      report(label, r);
    }
    const t = performance.now();
    const n = await trackedChanges().then((x) => x.length);
    report("read: count tracked changes", { ms: Math.round(performance.now() - t), trackedChanges: n });
  }

  if (cfg.steps?.includes("floor")) await floor();

  const s = await readStore();
  report("stored data", { parts: s.parts, bytes: s.bytes, citations: Object.keys(s.store.citations).length });
}

/**
 * Where Word sets the floor, in minimal experiments outside the pane: the
 * same 100 citation writes through Office.js and through Word's object
 * model, how the cost of a write grows in one session, and the ways to write
 * a bibliography. Only with --steps floor; it changes the document.
 */
async function floor() {
  const cites = (c) => citations(c);
  const officeJs = async (name, from, n = 100) => {
    const r = await word(async (c) => {
      const a = (await cites(c)).slice(from, from + n);
      const t = performance.now();
      a.forEach((x, i) => x.insertText(i % 2 ? "[x]" : "[y]", "Replace"));
      await c.sync();
      return { msPerWrite: +((performance.now() - t) / a.length).toFixed(1) };
    });
    report(name, r);
    return r;
  };
  await officeJs("floor: 100 citation writes, Office.js", 0);
  const com = await command("$cs = $doc.ContentControls; $sw = [Diagnostics.Stopwatch]::StartNew(); for ($i = 201; $i -le 300; $i++) { $cs.Item($i).Range.Text = '[c]' }; $sw.ElapsedMilliseconds");
  report("floor: 100 citation writes, Word's object model", { msPerWrite: +(com / 100).toFixed(1) });
  await officeJs("floor: the same 100 writes again, the text unchanged", 0);
  for (let k = 1; k <= 10; k++) await officeJs(`floor: round ${k} of 200 writes in one session`, 400 + ((k * 200) % 3000), 200);
  await command("$doc.UndoClear(); 1");
  await officeJs("floor: 100 writes after clearing the undo list", 0);
  const reviewed = await word(async (c) => {
    const a = (await cites(c)).slice(0, 50);
    const t = performance.now();
    a.forEach((x) => x.getReviewedText("Current"));
    await c.sync();
    return { msPerControl: +((performance.now() - t) / 50).toFixed(1) };
  });
  report("floor: getReviewedText, per control", reviewed);
  const ooxml = await word(async (c) => {
    const t = performance.now();
    const x = c.document.body.getOoxml();
    await c.sync();
    return { ms: Math.round(performance.now() - t), bytes: x.value.length };
  });
  report("floor: body.getOoxml", ooxml);
  const bib = async (name, write) => {
    const r = await word(async (c) => {
      const all = c.document.body.contentControls;
      all.load("items/tag");
      await c.sync();
      const b = all.items.find(isBib);
      const ps = b.paragraphs;
      ps.load("items/text");
      await c.sync();
      const t = performance.now();
      const n = write(b, ps.items);
      await c.sync();
      return { ms: Math.round(performance.now() - t), entries: n };
    });
    report(name, r);
  };
  await bib("floor: bibliography, 20 entries rewritten one by one", (b, ps) => { ps.slice(0, 20).forEach((p) => p.insertText(p.text + " ", "Replace")); return 20; });
  await bib("floor: bibliography, whole, one OOXML insert", (b, ps) => {
    const esc = (x) => x.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const body = ps.map((p) => `<w:p><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr><w:r><w:t xml:space="preserve">${esc(p.text)}</w:t></w:r></w:p>`).join("");
    b.getRange("Content").insertOoxml(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`, "Replace");
    return ps.length;
  });
  await bib("floor: bibliography, whole, as main wrote it (a write per run and per indent)", (b, ps) => {
    const texts = ps.map((p) => p.text);
    texts.forEach((x, i) => {
      const p = i ? b.insertParagraph("", "End") : b.paragraphs.getFirst();
      const [num, rest] = x.split("\t");
      for (const [k, part] of [num, "\t", rest ?? ""].entries()) {
        const r = i === 0 && k === 0 ? b.insertText(part, "Replace") : p.insertText(part, "End");
        r.font.italic = false; r.font.bold = false; r.font.superscript = false; r.font.subscript = false;
      }
      p.leftIndent = 28.35;
      p.firstLineIndent = -28.35;
    });
    return texts.length;
  });
}

onReady(async () => {
  let cfg = {};
  try { cfg = await (await fetch("measure-config.json", { cache: "no-store" })).json(); } catch { /* every step */ }
  await scenario(cfg);
});
