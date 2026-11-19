import { test } from "node:test";
import assert from "node:assert/strict";
import { refresh, readControls } from "../src/word.js";
import { fastEligible, FAST_CITES_MIN } from "../src/fast.js";
import { controlTexts } from "../src/ooxml.js";
import { CITE_PREFIX, BIB_TAG } from "../src/model.js";
import { markup, bodyPackage } from "./ooxml-fixtures.js";

const { sdt, r, p } = markup();
const t = (n) => CITE_PREFIX + n;

const renderer = {
  render: (secs) => secs.map((s) => ({ citations: s.citations.map((_, i) => [{ text: `[${i + 1}]` }]), entries: s.citations.map((c) => [{ text: c.items[0].id }]) })),
};

const session = { renderer, locale: "en-US" };

function storeFor(n, { drop = [] } = {}) {
  const citations = {};
  for (let i = 1; i <= n; i++) if (!drop.includes(String(i))) citations[String(i)] = [{ id: "a" }];
  return { sources: { a: { id: "a" } }, citations, settings: { locale: "en-US" } };
}

/**
 * A document of `n` stale citations, like read.test.js's context but with a
 * live text backing, a track-changes mode, and a body OOXML the fast path
 * can read and write. `insertOoxml` applies the package through the real
 * parser, like Word would; `corrupt` names a tag it leaves behind, and
 * `failInsert` makes the write throw.
 */
function context(n, { mode = "Off", tracked = 0, bookmark = "", corrupt = null, failInsert = false, bibParas = 0, emptyFirst = false } = {}) {
  const backing = Array.from({ length: n }, () => "[0]");
  if (emptyFirst) backing[0] = "";
  const bibText = bibParas ? Array.from({ length: bibParas }, (_, k) => `e${k + 1}`).join("\r") : null;
  if (bibText !== null) backing.push(bibText);
  const perControlInsert = [];
  const insertedPackages = [];
  const deleted = [];
  const bibWrites = [];
  const order = [];
  let queue = [];
  const later = (f) => queue.push(f);
  const items = backing.map((_, i) => ({
    tag: i < n ? t(i + 1) : BIB_TAG, title: "", id: i, placeholderText: "",
    get text() { return backing[i]; },
    load() {},
    delete() { deleted.push(i); },
    insertText(text) { perControlInsert.push(i); backing[i] = text; return {}; },
  }));
  const byId = (id) => items.find((c) => c.id === id);
  const bibParagraphs = () => ({
    load() {},
    items: [],
    getFirst() {
      const x = {};
      x.getOoxml = () => { const y = {}; later(() => { y.value = bodyPackage(p(r("b"))); }); return y; };
      return x;
    },
  });
  const currentPackage = () => bodyPackage(backing.slice(0, n).map((text, i) => p(
    (i === 0 && bookmark ? `<w:bookmarkStart w:id="0" w:name="${bookmark}"/><w:bookmarkEnd w:id="0"/>` : "")
    + sdt(t(i + 1), r(text)),
  )).join("") + (bibText !== null
    ? sdt(BIB_TAG, Array.from({ length: bibParas }, (_, k) => p(r(`e${k + 1}`))).join(""))
    : ""));
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
            if (found.tag === BIB_TAG) {
              return {
                isNullObject: false, load() {}, paragraphs: bibParagraphs(),
                getRange: () => ({ insertOoxml: (xml) => { bibWrites.push(xml); order.push("bib"); } }),
              };
            }
            return { isNullObject: false, load() {}, insertText: (text) => found.insertText(text) };
          },
        },
        getTrackedChanges() {
          const c = { load() { if (tracked === "move") throw new Error("GeneralException"); else later(() => { c.items = Array(tracked).fill({}); }); } };
          return c;
        },
        getOoxml() { const x = {}; later(() => { x.value = currentPackage(); }); return x; },
        insertOoxml(xml) {
          order.push("body");
          insertedPackages.push(xml);
          if (failInsert) throw new Error("Word is busy");
          for (const c of controlTexts(xml)) {
            const i = items.findIndex((x) => x.tag === c.tag);
            if (i >= 0 && items[i].tag.startsWith(CITE_PREFIX) && c.tag !== corrupt) backing[i] = c.current;
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
  return { ctx, backing, perControlInsert, insertedPackages, deleted, bibWrites, order };
}

test("fastEligible: only untracked, cites-only, large renumbers", () => {
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 0, cites: FAST_CITES_MIN }), true);
  assert.equal(fastEligible({ tracked: true, bibs: 0, removes: 0, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 1, removes: 0, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 1, cites: 10000 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 0, removes: 0, cites: FAST_CITES_MIN - 1 }), false);
  assert.equal(fastEligible({ tracked: false, bibs: 2, removes: 3, cites: 10000, withOthers: true }), true);
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
  assert.equal(r.fast, true);
  assert.deepEqual(backing, backing.map((_, i) => `[${i + 1}]`));
  assert.equal(r.cited, n);
});

test("track changes on, or the mode not off, keeps the per-control path", async () => {
  const tracked = context(3, { tracked: 1 });
  const r1 = await refresh(tracked.ctx, session, { store: storeFor(3) });
  assert.equal(tracked.insertedPackages.length, 0);
  assert.equal(tracked.perControlInsert.length, 3);
  assert.equal(r1.work, null);
  assert.equal(r1.fast, false);

  const coauthored = context(FAST_CITES_MIN, { mode: "TrackAll" });
  const r2 = await refresh(coauthored.ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(coauthored.insertedPackages.length, 0);
  assert.equal(r2.work.citations.length, FAST_CITES_MIN - 30);
  assert.equal(r2.work.total, FAST_CITES_MIN);
  assert.equal(r2.work.done, 30);
  assert.equal(r2.fast, false);
  assert.equal(coauthored.ctx.document.changeTrackingMode, "TrackAll");
});

test("a person's bookmark keeps the per-control path", async () => {
  const { ctx, insertedPackages, perControlInsert } = context(FAST_CITES_MIN, { bookmark: "keepMe" });
  const r = await refresh(ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(insertedPackages.length, 0);
  assert.equal(perControlInsert.length, 30);
  assert.equal(r.fast, false);
  assert.equal(r.work.citations.length, FAST_CITES_MIN - 30);
});

test("a Word-internal bookmark does not block the fast path", async () => {
  const { ctx, insertedPackages, perControlInsert } = context(FAST_CITES_MIN, { bookmark: "_GoBack" });
  const r = await refresh(ctx, session, { store: storeFor(FAST_CITES_MIN) });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, []);
  assert.equal(r.work, null);
  assert.equal(r.fast, true);
});

test("a fresh citation writes per control, the rest go in one body write", async () => {
  const n = FAST_CITES_MIN + 1;
  const { ctx, backing, perControlInsert, insertedPackages } = context(n);
  const r = await refresh(ctx, session, { store: storeFor(n), fresh: [0] });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, [0]);
  assert.equal(r.work, null);
  assert.equal(r.fast, true);
  assert.deepEqual(backing, backing.map((_, i) => `[${i + 1}]`));
});

test("a control the body write leaves behind is healed per control", async () => {
  const n = FAST_CITES_MIN;
  const { ctx, backing, perControlInsert, insertedPackages } = context(n, { corrupt: t(7) });
  const r = await refresh(ctx, session, { store: storeFor(n) });
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, [6]);
  assert.equal(r.work, null);
  assert.equal(r.fast, true);
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
  assert.equal(r.fast, false);
});

test("an emptied control is deleted first, then the rest go in one body write", async () => {
  const n = FAST_CITES_MIN + 1;
  const { ctx, backing, deleted, perControlInsert, insertedPackages } = context(n, { emptyFirst: true });
  const r = await refresh(ctx, session, { store: storeFor(n, { drop: ["1"] }) });
  assert.deepEqual(deleted, [0]);
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(perControlInsert, []);
  assert.equal(r.work, null);
  assert.equal(r.fast, true);
  assert.deepEqual(backing.slice(1), backing.slice(1).map((_, k) => `[${k + 1}]`));
});

test("a rewritten bibliography goes first, then the rest go in one body write", async () => {
  const n = FAST_CITES_MIN;
  const { ctx, perControlInsert, insertedPackages, bibWrites, order } = context(n, { bibParas: 130 });
  const r = await refresh(ctx, session, { store: storeFor(n) });
  assert.equal(bibWrites.length, 1);
  assert.equal(insertedPackages.length, 1);
  assert.deepEqual(order, ["bib", "body"]);
  assert.deepEqual(perControlInsert, []);
  assert.equal(r.work, null);
  assert.equal(r.fast, true);
});

test("a refresh with nothing to write reports no body write", async () => {
  const { ctx, backing, insertedPackages, perControlInsert } = context(3);
  backing[0] = "[1]";
  backing[1] = "[2]";
  backing[2] = "[3]";
  const r = await refresh(ctx, session, { store: storeFor(3) });
  assert.equal(r.work, null);
  assert.equal(r.fast, false);
  assert.deepEqual(insertedPackages, []);
  assert.deepEqual(perControlInsert, []);
});
