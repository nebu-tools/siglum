// What a refresh writes: only what differs, as a pure function of what the
// document holds and the store.
import { test } from "node:test";
import assert from "node:assert/strict";
import { diffEntries, paragraphsOf, plan, CITE_PREFIX } from "../src/model.js";
import { runsText } from "../src/render.js";

// A renderer that numbers by first appearance and lists "n. key", like a numeric style.
const renderer = {
  calls: 0,
  render(secs) {
    this.calls++;
    return secs.map((s) => {
      const order = [];
      for (const c of s.citations) for (const x of c.items) if (!order.includes(x.id)) order.push(x.id);
      return {
        citations: s.citations.map((c) => [{ text: `[${c.items.map((x) => order.indexOf(x.id) + 1).join(",")}]` }]),
        entries: order.map((id, i) => [{ text: `${i + 1}.\t${id}` }]),
      };
    });
  },
};
const sources = Object.fromEntries("abcdefgh".split("").map((k) => [k, { id: k }]));
const cite = (id, current, extra = {}) => ({ kind: "cite", tag: CITE_PREFIX + id, title: "", current, original: current, ...extra });
const bib = (...entries) => ({ kind: "bib", current: entries.map((e) => e + "\r").join("") });
const store = (citations) => ({ sources, citations, settings: { locale: "en-US" } });

test("a bibliography's paragraphs from its text", () => {
  assert.deepEqual(paragraphsOf("1.\ta\r2.\tb\r"), ["1.\ta", "2.\tb"]);
  assert.deepEqual(paragraphsOf("one"), ["one"]);
  assert.deepEqual(paragraphsOf(""), [""]);
});

test("diffEntries: nothing, a change, an insertion, a removal", () => {
  assert.deepEqual(diffEntries(["a", "b"], ["a", "b"]), []);
  assert.deepEqual(diffEntries(["a", "b", "c"], ["a", "x", "c"]), [{ at: 1, remove: 1, insert: [1] }]);
  assert.deepEqual(diffEntries(["a", "c"], ["a", "b", "c"]), [{ at: 1, remove: 0, insert: [1] }]);
  assert.deepEqual(diffEntries(["b", "c"], ["a", "b", "c"]), [{ at: 0, remove: 0, insert: [0] }]);
  assert.deepEqual(diffEntries(["a", "b", "c"], ["a", "c"]), [{ at: 1, remove: 1, insert: [] }]);
  assert.deepEqual(diffEntries(["a", "b", "c", "d"], ["b", "x", "d", "e"]), [{ at: 0, remove: 1, insert: [] }, { at: 2, remove: 1, insert: [1] }, { at: 4, remove: 0, insert: [3] }]);
});

test("diffEntries: applying the hunks gives the wanted entries", () => {
  const rand = (() => { let s = 7; return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32); })();
  for (let round = 0; round < 300; round++) {
    const have = Array.from({ length: Math.floor(rand() * 30) }, () => "e" + Math.floor(rand() * 12));
    const want = have.filter(() => rand() > 0.15).map((x) => (rand() < 0.1 ? x + "'" : x));
    for (let k = 0; k < 3; k++) want.splice(Math.floor(rand() * (want.length + 1)), 0, "n" + k);
    const hunks = diffEntries(have, want);
    if (!hunks) continue;
    const out = [...have];
    for (const h of [...hunks].reverse()) out.splice(h.at, h.remove, ...h.insert.map((j) => want[j]));
    assert.deepEqual(out, want);
  }
});

test("diffEntries: a bibliography where most entries change is written whole", () => {
  const have = Array.from({ length: 200 }, (_, i) => `${i + 1}.\tsource ${i}`);
  const want = ["1.\tnew", ...Array.from({ length: 200 }, (_, i) => `${i + 2}.\tsource ${i}`)];
  assert.equal(diffEntries(have, want), null);
  assert.equal(diffEntries(have.slice(0, 10), want.slice(0, 11)).length, 1, "a short one is written by its entries");
});

test("a document that holds what the store says needs no writes", () => {
  const p = plan([cite("1", "[1]"), cite("2", "[1,2]"), bib("1.\ta", "2.\tb")], store({ 1: [{ id: "a" }], 2: [{ id: "a" }, { id: "b" }] }), renderer);
  assert.deepEqual([p.citations, p.bibs, p.removes, p.store], [[], [], [], null]);
});

test("only the citations whose number changed are written, and the bibliography by its entries", () => {
  // c is cited before b now, so b becomes 3 and c 2.
  const controls = [cite("1", "[1]"), cite("3", "[?]"), cite("2", "[2]"), cite("4", "[1]"), bib("1.\ta", "2.\tb", "3.\td"), cite("5", "[1]"), bib("1.\te")];
  const st = store({ 1: [{ id: "a" }], 2: [{ id: "b" }], 3: [{ id: "c" }], 4: [{ id: "a" }], 5: [{ id: "e" }] });
  const p = plan(controls, st, renderer);
  assert.deepEqual(p.citations.map((c) => [c.index, runsText(c.runs)]), [[1, "[2]"], [2, "[3]"]]);
  assert.equal(p.bibs.length, 1);
  assert.deepEqual(p.bibs[0].hunks, [{ at: 1, remove: 2, insert: [1, 2] }]);
  assert.deepEqual(p.bibs[0].entries.map(runsText), ["1.\ta", "2.\tc", "3.\tb"]);
});

test("stored citations whose control is gone leave the store; those waiting for review or elsewhere stay", () => {
  const st = store({ 1: [{ id: "a" }], 2: [{ id: "b" }], gone: [{ id: "c" }], pending: [{ id: "d" }], note: [{ id: "e" }] });
  const p = plan([cite("1", "[1]"), cite("2", "[2]"), cite("pending", "", { original: "[3]" })], st, renderer, { elsewhere: ["note"] });
  assert.deepEqual(Object.keys(p.store.citations).sort(), ["1", "2", "note", "pending"]);
  assert.deepEqual(plan([cite("1", "[1]")], store({ 1: [{ id: "a" }] }), renderer).store, null);
});

test("a citation emptied by hand goes, and so does its stored entry", () => {
  const p = plan([cite("1", "", { original: "" }), cite("2", "[1]")], store({ 1: [{ id: "a" }], 2: [{ id: "b" }] }), renderer);
  assert.deepEqual(p.removes, [0]);
  assert.deepEqual(Object.keys(p.store.citations), ["2"]);
});

test("a pasted citation joins the store by the keys in its title", () => {
  const p = plan([cite("x", "[7]", { title: "Cite: a; b" })], store({}), renderer);
  assert.deepEqual(p.store.citations, { x: [{ id: "a" }, { id: "b" }] });
  assert.deepEqual(p.citations.map((c) => runsText(c.runs)), ["[1,2]"]);
});

test("two copies of one citation share its stored entry", () => {
  const p = plan([cite("1", "[1]"), cite("1", "[1]")], store({ 1: [{ id: "a" }] }), renderer);
  assert.deepEqual([p.citations, p.store, p.cited], [[], null, 2]);
});
