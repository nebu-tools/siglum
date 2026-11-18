// Probe page for e2e/probe.mjs: does things a user does in Word and reports
// what came out, as {probe: name, data}. The mode comes from probe-config.json.
//
// Copy, cut and paste go through getOoxml/insertOoxml, which is what Word's
// clipboard carries for a content control; the system clipboard stays untouched.
/* global Word */
import { post, sleep, $, word, harness, nextDone, onReady, idle, pick, key, lastDone, firstDone } from "./pane.js";
import * as names from "./names.mjs";

const report = (probe, data) => post({ probe, data });

/** Citations and lists in the body, as [kind, text]. */
const body = () => harness(() => word(async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag,items/text");
  await c.sync();
  return all.items.filter((x) => names.isCiteTag(x.tag) || x.tag === names.BIB_TAG).map((x) => (x.tag === names.BIB_TAG ? ["list", x.text.split(/[\r\n]+/).filter(Boolean).map((l) => l.slice(0, 22))] : ["cite", x.text]));
}));

/** Citation controls Word reaches elsewhere: footnotes, endnotes, headers. */
const elsewhere = () => harness(() => word(async (c) => {
  const out = {};
  const grab = async (label, coll) => {
    try {
      coll.load("items/tag,items/text");
      await c.sync();
      out[label] = coll.items.filter((x) => names.isCiteTag(x.tag)).map((x) => x.text);
    } catch (e) { out[label] = "API: " + (e.message ?? e); }
  };
  await grab("header (sections[0].getHeader)", c.document.sections.getFirst().getHeader("Primary").contentControls);
  try {
    const fn = c.document.body.footnotes;
    fn.load("items");
    await c.sync();
    await grab("footnote (body.footnotes)", fn.items[0].body.contentControls);
  } catch (e) { out["footnote (body.footnotes)"] = "API: " + (e.message ?? e); }
  try {
    const en = c.document.body.endnotes;
    en.load("items");
    await c.sync();
    await grab("endnote (body.endnotes)", en.items[0].body.contentControls);
  } catch (e) { out["endnote (body.endnotes)"] = "API: " + (e.message ?? e); }
  return out;
}));

/** The whole OOXML of the citation whose tag ends in `label`. */
const citeOoxml = (label) => harness(() => word(async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag");
  await c.sync();
  const x = all.items.find((y) => names.isCiteTag(y.tag) && y.tag.endsWith(label));
  if (!x) throw new Error("no citation " + label);
  const ox = x.getRange("Whole").getOoxml();
  await c.sync();
  return ox.value;
}));

const deleteCite = (label) => harness(() => word(async (c) => {
  const all = c.document.body.contentControls;
  all.load("items/tag");
  await c.sync();
  // Like selecting the citation and pressing Delete: the control and its text go.
  all.items.find((y) => names.isCiteTag(y.tag) && y.tag.endsWith(label)).getRange("Whole").delete();
  await c.sync();
}));

const pasteAfter = (text, ooxml) => harness(() => word(async (c) => {
  const r = c.document.body.search(text, { matchCase: true });
  r.load("items");
  await c.sync();
  r.items[r.items.length - 1].getRange("End").insertOoxml(ooxml, "Replace");
  await c.sync();
}));

const cursorAfter = (text) => harness(() => word(async (c) => {
  const r = c.document.body.search(text, { matchCase: true });
  r.load("items");
  await c.sync();
  r.items[0].getRange("End").select("End");
  await c.sync();
}));

/** Moves the cursor and waits for the selection watcher; returns whether it refreshed. */
async function watcher(where) {
  const done = nextDone(8000).then(() => true, () => false);
  await cursorAfter(where);
  const fired = await done;
  await idle();
  return fired;
}

const store = () => harness(() => word(async (c) => {
  const p = c.document.customXmlParts.getByNamespace(names.NS);
  p.load("items");
  await c.sync();
  const x = p.items[0].getXml();
  await c.sync();
  return { parts: p.items.length, store: names.storeFromXml(x.value) };
}));

const tracked = () => harness(() => word(async (c) => {
  const t = c.document.body.getTrackedChanges();
  t.load("items/type,items/text");
  await c.sync();
  return t.items.map((x) => [x.type, x.text.slice(0, 30)]);
}));

async function places() {
  const d = firstDone;
  report("first refresh", { status: $("status").textContent, written: d && d.citations.length, unknown: d && d.unknown });
  report("body after first refresh", await body());
  report("seen through other collections", await elsewhere());

  // A five-work citation, whose title no longer fits.
  await cursorAfter("Body one");
  for (const k of ["alphaeins2020", "betazwei2019", "gammadrei2018", "deltavier2017"]) {
    $("q").value = k; $("q").dispatchEvent(new Event("input"));
    key($("q"), "Enter", { ctrlKey: true });
  }
  $("q").value = "epsilonfuenf2016"; $("q").dispatchEvent(new Event("input"));
  key($("q"), "Enter");
  await idle();
  const all = await harness(() => word(async (c) => {
    const cc = c.document.body.contentControls;
    cc.load("items/tag,items/title,items/text");
    await c.sync();
    return cc.items.filter((x) => names.isCiteTag(x.tag)).map((x) => [x.tag, x.title, x.text]);
  }));
  const multi = all.find((x) => x[1].endsWith("…") || x[2].includes(","));
  report("five-work citation", { title: multi?.[1], text: multi?.[2] });

  // Clips for the other document: short title, locator, long title.
  const clips = { short: await citeOoxml("A"), locator: await citeOoxml("B"), long: multi ? await citeOoxml(multi[0].slice(-6)) : null };
  await fetch("/log", { method: "POST", body: JSON.stringify({ clips }) });

  // Copy: B (section 1) pasted into section 2, then the cursor moves on.
  await pasteAfter("Ziel", clips.locator);
  report("copy B into section 2: watcher refreshed", await watcher("Ende"));
  report("after copy", await body());

  // Cut and move: C from the table in section 1 to section 2.
  const c = await citeOoxml("C");
  await deleteCite("C");
  await pasteAfter("Ziel", c);
  report("move C to section 2: watcher refreshed", await watcher("Body two"));
  report("after move", await body());

  // The same move with track changes on: S2's work goes to section 1.
  await harness(() => word(async (ctx) => { ctx.document.changeTrackingMode = "TrackAll"; await ctx.sync(); }));
  const s2 = await citeOoxml("S2");
  await deleteCite("S2");
  await pasteAfter("Body two", s2);
  const fired = await watcher("Ende");
  if (!fired) { $("refresh").click(); await idle(); }
  report("tracked move S2 to section 1", { watcher: fired, status: $("status").textContent, body: await body(), tracked: await tracked() });
  await harness(() => word(async (ctx) => { ctx.document.changeTrackingMode = "Off"; await ctx.sync(); }));

  const s = await store();
  report("store before saving", { parts: s.parts, citations: Object.keys(names.citationsOf(s.store)).length, works: Object.keys(names.libraryOf(s.store)).length });
}

async function reopen() {
  const d = firstDone;
  report("first refresh after reopening", { status: $("status").textContent, citations: d?.citations.length, bibs: d?.bibs.length, removes: d?.removes.length, pending: d?.pending });
  const s = await store();
  report("library after reopening", { parts: s.parts, works: Object.keys(names.libraryOf(s.store)).length, citations: Object.keys(names.citationsOf(s.store)).length });
  report("body after reopening", await body());
}

async function paste() {
  const clips = await (await fetch("probe-clips.json", { cache: "no-store" })).json();
  const where = { short: "Paste one", locator: "Paste two", long: "Paste three" };
  for (const [k, ox] of Object.entries(clips)) if (ox) await pasteAfter(where[k], ox);
  $("refresh").click();
  await idle();
  report("after pasting short, locator and long-title citations", { status: $("status").textContent, unknown: lastDone?.unknown, body: await body() });
  const s = await store();
  const cits = names.citationsOf(s.store);
  report("stored items of the pasted citations", Object.values(cits));
}

onReady(async () => {
  let cfg = { mode: "places" };
  try { cfg = await (await fetch("probe-config.json", { cache: "no-store" })).json(); } catch { /* default */ }
  let ok = true;
  try {
    await { places, reopen, paste }[cfg.mode]();
  } catch (e) {
    ok = false;
    report("error", String(e && (e.stack || e.message || e)));
  }
  await sleep(200);
  post({ done: true, ok });
});
void pick;
