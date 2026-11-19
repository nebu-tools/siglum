import test from "node:test";
import assert from "node:assert/strict";
import { readControls, refresh } from "../src/word.js";
import { OoxmlError } from "../src/ooxml.js";
import { CITE_PREFIX, BIB_TAG, plan } from "../src/model.js";
import { markup, bodyPackage } from "./ooxml-fixtures.js";

const { sdt, r, p } = markup();

/**
 * Content controls as Word lists them, with `text` as Word shows it, and the
 * body's OOXML. `tracked` is the number of tracked changes, or "move" for a
 * document whose tracked changes Word cannot list. `raw` replaces the whole
 * package, as it came. `writes` records every
 * write and every touch of the store.
 */
function context(controls, { tracked = 0, ooxml = "", raw, failSyncAt = [] } = {}) {
  const asked = [];
  const writes = [];
  let queue = [];
  const later = (f) => queue.push(f);
  const items = controls.map((c, i) => ({
    tag: c.tag, title: "", id: i, text: c.text, placeholderText: c.placeholderText ?? "",
    load() {},
    delete() { writes.push(["delete", i]); },
    insertText() { writes.push(["insertText", i]); return {}; },
  }));
  let fail = false;
  let syncs = 0;
  const ctx = {
    document: {
      load() {},
      get customXmlParts() { writes.push(["customXmlParts"]); return {}; },
      get contentControls() { writes.push(["document.contentControls"]); return { items, load() {} }; },
      body: {
        contentControls: { items, load() {}, getByIdOrNullObject() { return { isNullObject: true, load() {} }; } },
        getTrackedChanges() {
          asked.push("trackedChanges");
          const t = { load() { if (tracked === "move") fail = true; else later(() => { t.items = Array(tracked).fill({}); }); } };
          return t;
        },
        getOoxml() { asked.push("ooxml"); const x = {}; later(() => { x.value = raw ?? bodyPackage(ooxml); }); return x; },
      },
    },
    async sync() {
      if (failSyncAt.includes(++syncs)) throw new Error("Word is busy");
      const q = queue;
      queue = [];
      if (fail) { fail = false; throw new Error("GeneralException"); }
      q.forEach((f) => f());
    },
  };
  return { ctx, asked, writes };
}

const renderer = {
  render: (secs) => secs.map((s) => ({ citations: s.citations.map((_, i) => [{ text: `[${i + 1}]` }]), entries: s.citations.map((c) => [{ text: c.items[0].id }]) })),
};
const store = { sources: { a: { id: "a" }, b: { id: "b" } }, citations: { 1: [{ id: "a" }], 2: [{ id: "b" }], 3: [{ id: "b" }] }, settings: { locale: "en-US" } };
const t = (n) => CITE_PREFIX + n;
const deleted = (x) => `<w:del w:id="1" w:author="A"><w:r><w:delText>${x}</w:delText></w:r></w:del>`;

test("without tracked changes the texts are what Word lists, and nothing else is read", async () => {
  const { ctx, asked } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "" }, { tag: BIB_TAG, text: "a\rb\r" }]);
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, ["trackedChanges"]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", ""], ["a\rb\r", undefined]]);
  // A control that is really empty is a leftover, as before.
  assert.deepEqual(plan(controls, store, renderer).removes, [1]);
});

test("with tracked changes the texts come from the OOXML, whatever Word shows", async () => {
  // "All Markup" keeps the deleted "[2]" in the control's text.
  const ooxml = p(sdt(t(1), r("[1]")) + sdt(t(2), deleted("[2]")));
  const { ctx, asked } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }], { tracked: 1, ooxml });
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, ["trackedChanges", "ooxml"]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", "[2]"]]);
  assert.equal(plan(controls, store, renderer).pending, 1);
});

test("a citation moved with track changes on waits for review where it was, and is numbered where it went", async () => {
  // Word cannot list the tracked changes of a document with a move.
  const ooxml = p(sdt(t(1), r("[1]"))) + p(sdt(t(2), `<w:moveFrom w:id="2"><w:r><w:t>[2]</w:t></w:r></w:moveFrom>`))
    + p(sdt(t(3), r("[3]"))) + p(sdt(t(2), `<w:moveTo w:id="3"><w:r><w:t>[2]</w:t></w:r></w:moveTo>`));
  const { ctx } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }, { tag: t(3), text: "[3]" }, { tag: t(2), text: "[2]" }], { tracked: "move", ooxml });
  const { controls } = await readControls(ctx);
  const p2 = plan(controls, store, renderer);
  assert.equal(p2.pending, 1);
  assert.equal(p2.cited, 3);
  assert.deepEqual(p2.citations.map((c) => [c.index, c.runs[0].text]), [[2, "[2]"], [3, "[3]"]]);
});

test("citations in text boxes are left out on both sides, so the lists match", async () => {
  const boxed = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>${p(sdt(t(3), r("[?]")))}</w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent>${p(sdt(t(3), r("[?]")))}</w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  const { ctx } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }], { tracked: 1, ooxml: p(sdt(t(1), r("[1]")) + boxed + sdt(t(2), deleted("[2]"))) });
  const { controls } = await readControls(ctx);
  assert.deepEqual(controls.map((c) => [c.tag, c.current, c.original]), [[t(1), "[1]", "[1]"], [t(2), "", "[2]"]]);
});

test("when the OOXML does not hold the controls Word lists, the read refuses and says how many each has", async () => {
  const { ctx } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }], { tracked: 1, ooxml: p(sdt(t(1), r("[1]"))) });
  await assert.rejects(readControls(ctx), (e) => e instanceof OoxmlError && e.key === "error.ooxmlMismatch" && e.params.listed === 2 && e.params.read === 1);
  // The same count in another order refuses too.
  const swapped = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }], { tracked: 1, ooxml: p(sdt(t(2), r("[2]")) + sdt(t(1), r("[1]"))) });
  await assert.rejects(readControls(swapped.ctx), (e) => e instanceof OoxmlError && e.params.listed === 2 && e.params.read === 2);
});

test("unreadable OOXML: the refresh stops before the plan, nothing is removed and no stored entry is pruned", async () => {
  // Cut off right after a citation's w:sdtContent opens, so that a lenient
  // reader would see it empty both ways and the planner would remove it.
  const full = bodyPackage(p(sdt(t(1), r("[1]")) + sdt(t(2), r("[2]") + deleted("x"))));
  const cut = full.slice(0, full.lastIndexOf("<w:sdtContent>") + "<w:sdtContent>".length);
  const { ctx, writes } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]x" }], { tracked: 1, raw: cut });
  const session = { renderer, locale: "en-US" };
  await assert.rejects(refresh(ctx, session, { store }), (e) => e instanceof OoxmlError && e.key === "error.ooxml");
  assert.deepEqual(writes, []);
  assert.deepEqual(Object.keys(store.citations), ["1", "2", "3"]);
});

test("without tracked changes an emptied control reads as no text, like the reviewed texts", async () => {
  const { ctx, asked } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[empty citation]", placeholderText: "[empty citation]" }]);
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, ["trackedChanges"]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", ""]]);
  // A control that is really empty is a leftover, as before.
  assert.deepEqual(plan(controls, store, renderer).removes, [1]);
});

test("a failing first batch leaves the work for the queue instead of throwing the plan away", async () => {
  // Two citations the store knows, one of them stale: the plan writes one
  // batch, and Word fails its sync.
  const near = { sources: { a: { id: "a" }, b: { id: "b" } }, citations: { 1: [{ id: "a" }], 2: [{ id: "b" }] }, settings: { locale: "en-US" } };
  const { ctx } = context([{ tag: t(1), text: "[9]" }, { tag: t(2), text: "[2]" }], { failSyncAt: [5] });
  const session = { renderer, locale: "en-US" };
  const r = await refresh(ctx, session, { store: near });
  // The plan stands, with the whole batch still to write: the queue's
  // batches write it, and plan again when they fail too.
  assert.equal(r.cited, 2);
  assert.equal(r.work.total, 1);
  assert.equal(r.work.citations.length, 1);
});
