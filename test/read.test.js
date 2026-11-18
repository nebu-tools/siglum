import test from "node:test";
import assert from "node:assert/strict";
import { readControls } from "../src/word.js";
import { CITE_PREFIX, BIB_TAG, plan } from "../src/model.js";

const sdt = (tag, content) => `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const r = (t) => `<w:r><w:t>${t}</w:t></w:r>`;

/**
 * Content controls as Word lists them, with `text` as Word shows it, and the
 * body's OOXML. `tracked` is the number of tracked changes, or "move" for a
 * document whose tracked changes Word cannot list.
 */
function context(controls, { tracked = 0, ooxml = "" } = {}) {
  const asked = [];
  let queue = [];
  const later = (f) => queue.push(f);
  const items = controls.map((c, i) => ({
    tag: c.tag, title: "", id: i, text: c.text, placeholderText: "[empty citation]",
    getReviewedText(kind) { asked.push(kind); const x = {}; later(() => { x.value = c[kind]; }); return x; },
  }));
  let fail = false;
  const ctx = {
    document: {
      body: {
        contentControls: { items, load() {} },
        getTrackedChanges() {
          asked.push("trackedChanges");
          const t = { load() { if (tracked === "move") fail = true; else later(() => { t.items = Array(tracked).fill({}); }); } };
          return t;
        },
        getOoxml() { asked.push("ooxml"); const x = {}; later(() => { x.value = `<w:body>${ooxml}</w:body>`; }); return x; },
      },
    },
    async sync() {
      const q = queue;
      queue = [];
      if (fail) { fail = false; throw new Error("GeneralException"); }
      q.forEach((f) => f());
    },
  };
  return { ctx, asked };
}

const renderer = {
  render: (secs) => secs.map((s) => ({ citations: s.citations.map((_, i) => [{ text: `[${i + 1}]` }]), entries: s.citations.map((c) => [{ text: c.items[0].id }]) })),
};
const store = { sources: { a: { id: "a" }, b: { id: "b" } }, citations: { 1: [{ id: "a" }], 2: [{ id: "b" }], 3: [{ id: "b" }] }, settings: { locale: "en-US" } };
const t = (n) => CITE_PREFIX + n;

test("without tracked changes the texts are what Word lists, and nothing else is read", async () => {
  const { ctx, asked } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "" }, { tag: BIB_TAG, text: "a\rb\r" }]);
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, ["trackedChanges"]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", ""], ["a\rb\r", undefined]]);
  assert.deepEqual(plan(controls, store, renderer).removes, [1]);
});

test("with tracked changes the texts come from the OOXML, whatever Word shows", async () => {
  // "All Markup" keeps the deleted "[2]" in the control's text.
  const ooxml = `<w:p>${sdt(t(1), r("[1]"))}${sdt(t(2), `<w:del><w:r><w:delText>[2]</w:delText></w:r></w:del>`)}</w:p>`;
  const { ctx, asked } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }], { tracked: 1, ooxml });
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, ["trackedChanges", "ooxml"]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", "[2]"]]);
  assert.equal(plan(controls, store, renderer).pending, 1);
});

test("a citation moved with track changes on waits for review where it was, and is numbered where it went", async () => {
  // Word cannot list the tracked changes of a document with a move.
  const ooxml = `<w:p>${sdt(t(1), r("[1]"))}</w:p><w:p>${sdt(t(2), `<w:moveFrom><w:r><w:t>[2]</w:t></w:r></w:moveFrom>`)}</w:p>`
    + `<w:p>${sdt(t(3), r("[3]"))}</w:p><w:p>${sdt(t(2), `<w:moveTo><w:r><w:t>[2]</w:t></w:r></w:moveTo>`)}</w:p>`;
  const { ctx } = context([{ tag: t(1), text: "[1]" }, { tag: t(2), text: "[2]" }, { tag: t(3), text: "[3]" }, { tag: t(2), text: "[2]" }], { tracked: "move", ooxml });
  const { controls } = await readControls(ctx);
  const p = plan(controls, store, renderer);
  assert.equal(p.pending, 1);
  assert.equal(p.cited, 3);
  assert.deepEqual(p.citations.map((c) => [c.index, c.runs[0].text]), [[2, "[2]"], [3, "[3]"]]);
});

test("when the OOXML does not match the controls, each control's reviewed text is read", async () => {
  const { ctx, asked } = context([{ tag: t(1), text: "[1]", Current: "<[1]>", Original: "<[1]>" }, { tag: t(2), text: "[2]", Current: "<[empty citation]>", Original: "<[2]>" }], { tracked: 1, ooxml: sdt(t(1), r("[1]")) });
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked.filter((x) => x === "Current" || x === "Original").length, 4);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", "[2]"]]);
});
