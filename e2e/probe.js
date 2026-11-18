// The probe page for e2e/probe.mjs: does what a person does in Word and
// reports what came out, as {probe: name, data}. These are findings, not
// checks. The mode comes from probe-config.json.
import { post, $, word, onReady, done, refresh, controls, readStore, isCite } from "./pane.js";
import { CITE_PREFIX } from "../src/model.js";

const report = (probe, data) => post({ probe, data });

/** Citations and bibliographies in the body: citation texts, and each bibliography's first words per entry. */
const body = async () => (await controls()).map((x) => (x.kind === "bib" ? x.text.split(/[\r\n]+/).filter(Boolean).map((l) => l.slice(0, 22)) : x.text));

/** Citation controls in other stories, as far as Office.js reaches them. */
const elsewhere = () => word(async (c) => {
  const out = {};
  const grab = async (label, get) => {
    try {
      const coll = get();
      coll.load("items/tag,items/text");
      await c.sync();
      out[label] = coll.items.filter(isCite).map((x) => x.text);
    } catch (e) { out[label] = "API: " + (e.message ?? e); }
  };
  await grab("header", () => c.document.sections.getFirst().getHeader("Primary").contentControls);
  for (const kind of ["footnotes", "endnotes"]) {
    try {
      const notes = c.document.body[kind];
      notes.load("items");
      await c.sync();
      await grab(kind, () => notes.items[0].body.contentControls);
    } catch (e) { out[kind] = "API: " + (e.message ?? e); }
  }
  return out;
});

/** The OOXML of the citation whose tag ends in `label`, with the control around it. */
const citeOoxml = (label) => word(async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag,items/title");
  await c.sync();
  const x = all.items.find((y) => isCite(y) && y.tag.endsWith(label));
  const ox = x.getRange("Whole").getOoxml();
  await c.sync();
  return ox.value;
});

const pasteAfter = (text, ooxml) => word(async (c) => {
  const r = c.document.body.search(text, { matchCase: true });
  r.load("items");
  await c.sync();
  r.items[r.items.length - 1].getRange("End").insertOoxml(ooxml, "Replace");
  await c.sync();
});

async function places() {
  const d = done.first;
  report("first refresh", { status: $("status").textContent, written: d?.citations.length, unknown: d?.unknown });
  report("body after the first refresh", await body());
  report("citations in other stories, read through Office.js", await elsewhere());
  // A citation of five sources with long keys, whose title is cut at 64 characters.
  const long = await word(async (c) => {
    const all = c.document.body.contentControls;
    all.load("items/tag,items/title,items/text");
    await c.sync();
    const x = all.items.find((y) => isCite(y) && y.tag.endsWith("LONG"));
    return { title: x.title, text: x.text };
  });
  report("five-source citation", long);
  await post({ clips: { short: await citeOoxml("A"), page: await citeOoxml("B"), long: await citeOoxml("LONG") } });
  const s = await readStore();
  report("store", { parts: s.parts, citations: Object.keys(s.store.citations).length, sources: Object.keys(s.store.sources).length });
}

async function paste() {
  const clips = await (await fetch("probe-clips.json", { cache: "no-store" })).json();
  const where = { short: "Paste one", page: "Paste two", long: "Paste three" };
  for (const [k, ox] of Object.entries(clips)) await pasteAfter(where[k], ox);
  const r = await refresh();
  report("after pasting a short, a paged and a cut-title citation", { status: $("status").textContent, unknown: r?.unknown, body: await body() });
  const s = await readStore();
  report("stored items of the pasted citations", Object.entries(s.store.citations).map(([id, items]) => [id.slice(-6), items]));
}

async function reopen() {
  const d = done.first;
  report("first refresh after reopening", { status: $("status").textContent, citations: d?.citations.length, bibs: d?.bibs.length, removes: d?.removes.length });
  report("body after reopening", await body());
  const s = await readStore();
  const tags = new Set((await controls()).filter((x) => x.kind === "cite").map((x) => x.tag.slice(CITE_PREFIX.length)));
  report("stored citations without a control in the body", Object.keys(s.store.citations).filter((t) => !tags.has(t)).length);
}

onReady(async () => {
  const cfg = await (await fetch("probe-config.json", { cache: "no-store" })).json();
  await { places, paste, reopen }[cfg.mode]();
});
