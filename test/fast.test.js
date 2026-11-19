import { test } from "node:test";
import assert from "node:assert/strict";
import { refresh, readControls, fastEligible, FAST_CITES_MIN } from "../src/word.js";
import { controlTexts } from "../src/ooxml.js";
import { CITE_PREFIX } from "../src/model.js";
import { markup, bodyPackage } from "./ooxml-fixtures.js";

const { sdt, r, p } = markup();
const t = (n) => CITE_PREFIX + n;

const renderer = {
  render: (secs) => secs.map((s) => ({ citations: s.citations.map((_, i) => [{ text: `[${i + 1}]` }]), entries: s.citations.map((c) => [{ text: c.items[0].id }]) })),
};

const session = { renderer, locale: "en-US" };

function storeFor(n) {
  const citations = {};
  for (let i = 1; i <= n; i++) citations[String(i)] = [{ id: "a" }];
  return { sources: { a: { id: "a" } }, citations, settings: { locale: "en-US" } };
}

/**
 * A document of `n` stale citations, like read.test.js's context but with a
 * live text backing, a track-changes mode, and a body OOXML the fast path
 * can read and write. `insertOoxml` applies the package through the real
 * parser, like Word would; `corrupt` names a tag it leaves behind, and
 * `failInsert` makes the write throw.
 */
function context(n, { mode = "Off", tracked = 0, bookmark = "", corrupt = null, failInsert = false } = {}) {
  const backing = Array.from({ length: n }, () => "[0]");
  const perControlInsert = [];
  const insertedPackages = [];
  let queue = [];
  const later = (f) => queue.push(f);
  const items = backing.map((_, i) => ({
    tag: t(i + 1), title: "", id: i, placeholderText: "",
    get text() { return backing[i]; },
    load() {},
    delete() {},
    insertText(text) { perControlInsert.push(i); backing[i] = text; return {}; },
  }));
  const byId = (id) => items.find((c) => c.id === id);
  const currentPackage = () => bodyPackage(backing.map((text, i) => p(
    (i === 0 && bookmark ? `<w:bookmarkStart w:id="0" w:name="${bookmark}"/><w:bookmarkEnd w:id="0"/>` : "")
    + sdt(t(i + 1), r(text)),
  )).join(""));
  const ctx = {
    document: {
      changeTrackingMode: mode,
      load() {},
      get contentControls() { return { items: [], load() {} }; },
      body: {
        contentControls: {
          items, load() {},
          getByIdOrNullObject(id) {
            const found = byId(id);
            if (!found) return { isNullObject: true, load() {} };
            return { isNullObject: false, load() {}, insertText: (text) => found.insertText(text) };
          },
        },
        getTrackedChanges() {
          const c = { load() { if (tracked === "move") throw new Error("GeneralException"); else later(() => { c.items = Array(tracked).fill({}); }); } };
          return c;
        },
        getOoxml() { const x = {}; later(() => { x.value = currentPackage(); }); return x; },
        insertOoxml(xml) {
          insertedPackages.push(xml);
          if (failInsert) throw new Error("Word is busy");
          for (const c of controlTexts(xml)) {
            const i = items.findIndex((x) => x.tag === c.tag);
            if (i >= 0 && c.tag !== corrupt) backing[i] = c.current;
          }
        },
      },
    },
    async sync() {
      const q = queue;
      queue = [];
      q.forEach((f) => f());
    },
  };
  return { ctx, backing, perControlInsert, insertedPackages };
}

test("fastEligible: only untracked, cites-only, large renumbers", () => {
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 0, cites: FAST_CITES_MIN }), true);
  assert.equal(fastEligible({ tracked: true, bibs: 0, removes: 0, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 1, removes: 0, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 1, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 0, cites: FAST_CITES_MIN - 1 }), false);
});

test("readControls reports whether tracked changes exist", async () => {
  const plain = context(1);
  assert.equal((await readControls(plain.ctx)).tracked, false);
  const revved = context(1, { tracked: 1 });
  assert.equal((await readControls(revved.ctx)).tracked, true);
});

test("five hundred stale citations go in one body write, verified", async () => {
  const n = 500;
  const { ctx, backing, perControlInsert, insertedPackages } = context(n);
  const r = await refresh(ctx, session, { store: storeFor(n) });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, []);
  assert.equal(r.work, null);
  assert.deepEqual(backing, backing.map((_, i) => `[${i + 1}]`));
  assert.equal(r.cited, n);
});

test("track changes on, or the mode not off, keeps the per-control path", async () => {
  const tracked = context(3, { tracked: 1 });
  const r1 = await refresh(tracked.ctx, session, { store: storeFor(3) });
  assert.equal(tracked.insertedPackages.length, 0);
  assert.equal(tracked.perControlInsert.length, 3);
  assert.equal(r1.work, null);

  const coauthored = context(FAST_CITES_MIN, { mode: "TrackAll" });
  const r2 = await refresh(coauthored.ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(coauthored.insertedPackages.length, 0);
  assert.equal(r2.work.citations.length, FAST_CITES_MIN - 30);
  assert.equal(r2.work.total, FAST_CITES_MIN);
  assert.equal(r2.work.done, 30);
  assert.equal(coauthored.ctx.document.changeTrackingMode, "TrackAll");
});

test("a person's bookmark keeps the per-control path", async () => {
  const { ctx, insertedPackages, perControlInsert } = context(FAST_CITES_MIN, { bookmark: "keepMe" });
  const r = await refresh(ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(insertedPackages.length, 0);
  assert.equal(perControlInsert.length, 30);
  assert.equal(r.work.citations.length, FAST_CITES_MIN - 30);
});

test("a Word-internal bookmark does not block the fast path", async () => {
  const { ctx, insertedPackages, perControlInsert } = context(FAST_CITES_MIN, { bookmark: "_GoBack" });
  const r = await refresh(ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, []);
  assert.equal(r.work, null);
});

test("a fresh citation writes per control, the rest go in one body write", async () => {
  const n = FAST_CITES_MIN + 1;
  const { ctx, backing, perControlInsert, insertedPackages } = context(n);
  const r = await refresh(ctx, session, { store: storeFor(n), fresh: [0] });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, [0]);
  assert.equal(r.work, null);
  assert.deepEqual(backing, backing.map((_, i) => `[${i + 1}]`));
});

test("a control the body write leaves behind is healed per control", async () => {
  const n = FAST_CITES_MIN;
  const { ctx, backing, perControlInsert, insertedPackages } = context(n, { corrupt: t(7) });
  const r = await refresh(ctx, session, { store: storeFor(n) });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, [6]);
  assert.equal(r.work, null);
  assert.equal(backing[6], "[7]");
});

test("a failing body write falls back to batches", async () => {
  const n = FAST_CITES_MIN;
  const { ctx, perControlInsert, insertedPackages } = context(n, { failInsert: true });
  const r = await refresh(ctx, session, { store: storeFor(n) });
  assert.equal(insertedPackages.length, 1);
  assert.equal(perControlInsert.length, 30);
  assert.equal(r.work.citations.length, n - 30);
  assert.equal(r.work.total, n);
  assert.equal(r.work.done, 30);
});
