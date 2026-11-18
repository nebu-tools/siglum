import test from "node:test";
import assert from "node:assert/strict";
import { readControls } from "../src/word.js";
import { CITE_PREFIX, BIB_TAG, plan } from "../src/model.js";

// Content controls as Word reports them: `text` with tracked changes
// accepted, and getReviewedText("Original") with them rejected.
function context(controls) {
  const asked = [];
  let queue = [];
  const items = controls.map((c, i) => ({
    tag: c.tag, title: c.title ?? "", id: i, text: c.text, placeholderText: "[empty citation]",
    getReviewedText(kind) {
      asked.push([c.tag, kind]);
      const r = {};
      queue.push(() => { r.value = c.original; });
      return r;
    },
  }));
  const ctx = {
    document: { body: { contentControls: { items, load() {} } } },
    async sync() { const q = queue; queue = []; q.forEach((f) => f()); },
  };
  return { ctx, asked };
}

const renderer = {
  render: (secs) => secs.map((s) => ({ citations: s.citations.map((_, i) => [{ text: `[${i + 1}]` }]), entries: s.citations.map((c) => [{ text: c.items[0].id }]) })),
};
const store = { sources: { a: { id: "a" }, b: { id: "b" } }, citations: { 1: [{ id: "a" }], 2: [{ id: "b" }], 3: [{ id: "b" }] }, settings: { locale: "en-US" } };

test("only a citation that reads empty is asked for its original text", async () => {
  const { ctx, asked } = context([
    { tag: CITE_PREFIX + "1", text: "[1]" },
    { tag: CITE_PREFIX + "2", text: "", original: "<[2]>" },
    { tag: BIB_TAG, text: "a\rb" },
  ]);
  const { controls } = await readControls(ctx);
  assert.deepEqual(asked, [[CITE_PREFIX + "2", "Original"]]);
  assert.deepEqual(controls.map((c) => [c.current, c.original]), [["[1]", "[1]"], ["", "[2]"], ["a\rb", undefined]]);
});

test("a citation moved away with track changes on waits for review, and its copy is numbered where it went", async () => {
  // Word gives a moved-away control an empty text; its reviewed text is the same both ways.
  const { ctx } = context([
    { tag: CITE_PREFIX + "1", text: "[1]" },
    { tag: CITE_PREFIX + "2", text: "", original: "<[2]>" },
    { tag: CITE_PREFIX + "3", text: "[3]" },
    { tag: CITE_PREFIX + "2", text: "[2]" },
  ]);
  const { controls } = await readControls(ctx);
  const p = plan(controls, store, renderer);
  assert.equal(p.pending, 1);
  assert.equal(p.cited, 3);
  assert.deepEqual(p.citations.map((c) => [c.index, c.runs[0].text]), [[2, "[2]"], [3, "[3]"]]);
});

test("a control emptied by hand shows its placeholder and goes", async () => {
  const { ctx } = context([{ tag: CITE_PREFIX + "1", text: "", original: "<[empty citation]>" }]);
  const { controls } = await readControls(ctx);
  assert.deepEqual(plan(controls, store, renderer).removes, [0]);
});
