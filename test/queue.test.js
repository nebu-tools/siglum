import { test } from "node:test";
import assert from "node:assert/strict";
import { createQueue } from "../src/queue.js";

const CHUNK = 30;

/**
 * A document of numbered citations and what the queue needs from Word. The
 * right text of a citation is its position, "[1]" for the first. refresh
 * plans like word.refresh: it writes the first CHUNK citations that differ
 * and leaves the rest as work. The test can hold the next Word call of a kind
 * ("write", "read", "signature"), or make it fail once.
 *
 * `log` records reads and writes with the plan they belong to, and the
 * queue's "settled" and "idle"; `stale` counts writes of a plan older than
 * the last one read.
 */
function fakeDocument(n) {
  const doc = { cites: Array.from({ length: n }, (_, i) => ({ id: "c" + i, text: "[0]" })), log: [], stale: 0, plans: 0 };
  const holds = new Map();
  const failures = new Set();
  doc.hold = (kind) => {
    let release;
    let reached;
    const gate = new Promise((r) => { release = r; });
    const at = new Promise((r) => { reached = r; });
    holds.set(kind, { gate, reached });
    return { reached: at, release };
  };
  doc.failNext = (kind) => failures.add(kind);
  const enter = async (kind) => {
    const h = holds.get(kind);
    if (h) { holds.delete(kind); h.reached(); await h.gate; }
    if (failures.delete(kind)) throw new Error(`${kind} failed`);
  };
  const write = (c, plan) => {
    if (plan !== doc.plans) doc.stale++;
    const cite = doc.cites.find((x) => x.id === c.id);
    if (cite) cite.text = c.text;
    doc.log.push(["write", plan]);
  };
  doc.right = () => doc.cites.every((c, i) => c.text === `[${i + 1}]`);
  doc.signature = async () => { await enter("signature"); return doc.cites.map((c) => c.id).join("|"); };
  doc.refresh = async () => {
    await enter("read");
    const plan = ++doc.plans;
    doc.log.push(["read", plan]);
    const differ = doc.cites.map((c, i) => ({ id: c.id, text: `[${i + 1}]`, plan })).filter((c, i) => doc.cites[i].text !== c.text);
    differ.slice(0, CHUNK).forEach((c) => write(c, plan));
    const rest = differ.slice(CHUNK);
    const work = rest.length ? { citations: rest, total: rest.length, done: 0 } : null;
    return { cited: doc.cites.length, work, signature: doc.cites.map((c) => c.id).join("|"), store: { sources: {} } };
  };
  doc.writeChunk = async (context, work, size) => {
    await enter("write");
    const batch = work.citations.slice(0, size);
    batch.forEach((c) => write(c, c.plan));
    work.citations = work.citations.slice(batch.length);
    work.done += batch.length;
  };
  return doc;
}

function setup(n = 150) {
  const doc = fakeDocument(n);
  let tick = 0;
  const changes = [];
  let idle = [];
  const q = createQueue({
    word: (fn) => fn({}),
    refresh: () => doc.refresh(),
    writeChunk: doc.writeChunk,
    signature: () => doc.signature(),
    chunk: CHUNK,
    // Every batch takes half a second, so the batches stay at CHUNK.
    now: () => (tick += 500),
    emit(type, detail) {
      if (type === "settled") doc.log.push(["settled"]);
      if (type === "change") {
        changes.push({ ...detail, progress: detail.progress && { ...detail.progress } });
        if (!detail.busy && changes.at(-2)?.busy !== false) { doc.log.push(["idle"]); idle.splice(0).forEach((f) => f()); }
      }
    },
  });
  const quiet = () => new Promise((r) => { if (!q.state.busy) r(); else idle.push(r); });
  const refresh = () => q.run(() => doc.refresh());
  return { doc, q, quiet, refresh, changes };
}

/** Starts a renumbering of 150 citations and holds its second batch. */
async function renumbering() {
  const s = setup();
  const r = await s.refresh();
  assert.equal(r.work.citations.length, 120);
  const h = s.doc.hold("write");
  await h.reached;
  return { ...s, release: h.release };
}

function assertEnd({ doc, q }, { settled = 1 } = {}) {
  assert.ok(doc.right(), JSON.stringify(doc.cites.filter((c, i) => c.text !== `[${i + 1}]`).slice(0, 3)));
  assert.equal(doc.stale, 0, "a batch of an old plan was written after a newer plan was read");
  const types = doc.log.map((e) => e[0]);
  assert.equal(types.filter((x) => x === "settled").length, settled);
  assert.equal(types.filter((x) => x === "idle").length, 1, "the queue went idle before the end");
  if (settled) assert.deepEqual(types.slice(-2), ["settled", "idle"]);
  assert.equal(q.state.busy, false);
  assert.equal(q.state.progress, null);
}

const importing = () => async () => ({ added: 2, updated: 0, library: {}, store: { sources: { a: {} } } });

test("a renumbering on its own: batches, then settled once", async () => {
  const s = setup();
  await s.refresh();
  await s.quiet();
  assertEnd(s);
  assert.deepEqual(s.doc.log.filter((e) => e[0] === "read"), [["read", 1]]);
  const progress = s.changes.map((c) => c.progress).filter(Boolean);
  assert.deepEqual(progress.map((x) => x.done), [0, 30, 60, 90]);
});

test("an import during the batches: it runs, the queue plans again and finishes, the progress stays meanwhile", async () => {
  const s = await renumbering();
  const imported = s.q.run(importing());
  s.release();
  const r = await imported;
  assert.equal(r.added, 2);
  // The import is done, the renumbering is not: the pane still has the progress to show.
  assert.ok(s.q.state.progress, "the progress was dropped with the import");
  assert.equal(s.doc.log.filter((e) => e[0] === "settled").length, 0);
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.plans, 2);
});

test("a second refresh during the batches takes over, and no batch of the first is written after it", async () => {
  const s = await renumbering();
  const second = s.refresh();
  s.release();
  assert.ok((await second).work);
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.plans, 2);
});

test("a new citation during the batches: its plan takes over with the new numbers", async () => {
  const s = await renumbering();
  const cited = s.q.run(() => { s.doc.cites.unshift({ id: "new", text: "[?]" }); return s.doc.refresh(); });
  s.release();
  await cited;
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.cites[0].text, "[1]");
});

test("the operation that took over fails: its error shows, the queue plans again and finishes", async () => {
  const s = await renumbering();
  const failed = s.q.run(async () => { throw new Error("boom"); });
  s.release();
  assert.equal(await failed, null);
  await s.quiet();
  assertEnd(s);
  assert.equal(s.q.state.error.message, "boom");
  assert.equal(s.q.state.stale, false);
});

test("the operation that took over fails, and so does planning again: stale until a refresh succeeds", async () => {
  const s = await renumbering();
  s.doc.failNext("read");
  const failed = s.q.run(async () => { throw new Error("boom"); });
  s.release();
  await failed;
  await s.quiet();
  assert.equal(s.q.state.stale, true);
  assert.equal(s.q.state.progress, null);
  assert.equal(s.q.state.error.message, "boom");
  assert.ok(!s.doc.right());
  assert.equal(s.doc.log.filter((e) => e[0] === "settled").length, 0);
  // The watcher waits; a click on Refresh continues.
  await s.q.check();
  assert.equal(s.doc.plans, 1);
  await s.refresh();
  assert.equal(s.q.state.stale, false);
  assert.equal(s.q.state.error, null);
  await s.quiet();
  assert.ok(s.doc.right());
  assert.equal(s.doc.stale, 0);
  assert.deepEqual(s.doc.log.filter((e) => e[0] === "settled").length, 1);
  assert.deepEqual(s.doc.log.at(-2), ["settled"]);
});

test("a batch fails: the queue plans again and finishes", async () => {
  const s = await renumbering();
  s.doc.failNext("write");
  s.release();
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.plans, 2);
});

test("a batch fails and planning again fails: stale, no settled", async () => {
  const s = await renumbering();
  s.doc.failNext("write");
  s.doc.failNext("read");
  s.release();
  await s.quiet();
  assert.equal(s.q.state.stale, true);
  assert.equal(s.q.state.error.message, "read failed");
  assert.equal(s.doc.log.filter((e) => e[0] === "settled").length, 0);
  // An import after that plans again too, and finishes the work.
  await s.q.run(importing());
  await s.quiet();
  assert.ok(s.doc.right());
  assert.equal(s.q.state.stale, false);
  assert.equal(s.doc.log.filter((e) => e[0] === "settled").length, 1);
});

test("two operations waiting: both run in order, then one plan", async () => {
  const s = await renumbering();
  const order = [];
  const a = s.q.run(async () => { order.push("a"); return { added: 1 }; });
  const b = s.q.run(async () => { order.push("b"); return { added: 2 }; });
  s.release();
  assert.deepEqual([(await a).added, (await b).added], [1, 2]);
  await s.quiet();
  assertEnd(s);
  assert.deepEqual(order, ["a", "b"]);
  assert.equal(s.doc.plans, 2);
});

test("a plan waiting behind an import: the plan is the one that finishes, no extra plan", async () => {
  const s = await renumbering();
  s.q.run(importing());
  s.refresh();
  s.release();
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.plans, 2);
});

test("the order of the controls changes during the batches: the queue plans again before it says settled", async () => {
  const s = await renumbering();
  // A citation cut and pasted by hand, without the pane.
  const [moved] = s.doc.cites.splice(140, 1);
  s.doc.cites.splice(3, 0, moved);
  s.release();
  await s.quiet();
  assertEnd(s);
  assert.equal(s.doc.plans, 2);
});

test("an operation without a plan on a current document neither plans nor says settled", async () => {
  const s = setup(10);
  await s.refresh();
  await s.quiet();
  await s.q.run(importing());
  await s.quiet();
  assert.equal(s.doc.plans, 1);
  assert.equal(s.doc.log.filter((e) => e[0] === "settled").length, 1);
});

test("the watcher plans only when the order changed since the document was current", async () => {
  const s = setup(10);
  await s.refresh();
  await s.quiet();
  await s.q.check();
  assert.equal(s.doc.plans, 1);
  s.doc.cites.pop();
  await s.q.check();
  await s.quiet();
  assert.equal(s.doc.plans, 2);
  assert.ok(s.doc.right());
});
