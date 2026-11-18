// The binding experiment's page: runs in the pane next to the real one and
// measures what e2e/experiments/bind.mjs asks for (bind-config.json). Every
// result goes to the harness as {k: name, v: value}.
/* global Word, Office */
import { post, word, command, onReady, sleep } from "../pane.js";
import { NS, ITEM_ID, TAG, BIB, esc, storeXml, richPackage, runXml, xpathFor } from "./shared.js";

const log = (k, v) => post({ k, v });
const T = () => performance.now();
const ms = (t) => Math.round(T() - t);
const err = (e) => String(e?.message ?? e) + (e?.code ? ` (${e.code})` : "") + (e?.debugInfo?.errorLocation ? ` at ${e.debugInfo.errorLocation}` : "");
const ps = (s) => `'${String(s).replace(/'/g, "''")}'`;

/** Runs `fn` and logs {ms, ...result}, or the error. */
async function step(name, fn) {
  const t = T();
  try {
    const v = await fn();
    log(name, { ms: ms(t), ...(v && typeof v === "object" && !Array.isArray(v) ? v : { value: v }) });
    return v;
  } catch (e) {
    log(name, { ms: ms(t), error: err(e) });
    return undefined;
  }
}

const sets = () => {
  const r = Office.context.requirements;
  const out = {};
  for (const [s, vs] of [["WordApi", ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]], ["WordApiDesktop", ["1.1", "1.2", "1.3", "1.4"]], ["WordApiHiddenDocument", ["1.5"]], ["CustomXmlParts", ["1.1"]]]) out[s] = vs.filter((v) => r.isSetSupported(s, v)).join(",");
  return out;
};

const part = (c) => {
  const p = c.document.customXmlParts.getByNamespace(NS);
  p.load("items/id");
  return p;
};
/** The part's values: id → text. */
async function partValues(c) {
  const p = part(c);
  await c.sync();
  const x = p.items[0].getXml();
  await c.sync();
  return new Map([...x.value.matchAll(/<c:c id="([^"]*)">([^<]*)<\/c:c>/g)].map((m) => [m[1], m[2].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&")]));
}

/** The body's controls in document order: [{tag, text}]. */
// body.contentControls leaves plain-text controls out; getContentControls() (WordApi 1.5) has them all.
async function bodyControls(c, more = "") {
  const all = c.document.body.getContentControls();
  all.load("items/tag,items/text" + more);
  await c.sync();
  return all.items;
}

// ---------- values and their expected texts ----------
let cfg;
let rich = false;
/** id → the element's value (text, or a package for rich controls). */
let values;
/** id → the text the control should show. */
let shown;
const richValue = (text) => richPackage(runXml({ text }));

/** New texts for some citations: "all", "tenth", "one", "none", or a count spread through the document. */
function change(fraction, round) {
  const ids = [...shown.keys()];
  const n = typeof fraction === "number" ? fraction : null;
  const pick = fraction === "all" ? ids : fraction === "tenth" ? ids.filter((_, i) => i % 10 === 0) : fraction === "one" ? [ids[Math.floor(ids.length / 2)]] : n ? ids.filter((_, i) => i % Math.floor(ids.length / n) === 0).slice(0, n) : [];
  const out = new Map();
  pick.forEach((id, k) => out.set(id, `[${round}.${k}]`));
  return out;
}

// ---------- the four ways to write the part ----------
const WRITERS = {
  // WordApi 1.4: replace the whole part.
  async setXml(c, changes) {
    for (const [id, t] of changes) { values.set(id, rich ? richValue(t) : t); shown.set(id, t); }
    const p = part(c);
    await c.sync();
    p.items[0].setXml(storeXml(values));
  },
  // WordApi 1.4: one call per changed element, by xpath.
  async updateElement(c, changes) {
    const p = part(c);
    await c.sync();
    for (const [id, t] of changes) {
      values.set(id, rich ? richValue(t) : t); shown.set(id, t);
      p.items[0].updateElement(xpathFor(id), `<c:c xmlns:c="${NS}" id="${esc(id)}">${esc(values.get(id))}</c:c>`, { c: NS });
    }
  },
  // WordApiDesktop 1.3: the node's text.
  async nodeText(c, changes) {
    const p = part(c);
    await c.sync();
    const nodes = p.items[0].selectNodes("/*/*");
    nodes.load("items/xpath");
    await c.sync();
    const at = new Map([...shown.keys()].map((id, i) => [id, i]));
    for (const [id, t] of changes) {
      values.set(id, rich ? richValue(t) : t); shown.set(id, t);
      nodes.items[at.get(id)].text = values.get(id);
    }
  },
  // WordApiDesktop 1.3: selectSingleNode by position for each changed element, no collection loaded.
  async nodeSingle(c, changes) {
    const p = part(c);
    await c.sync();
    const at = new Map([...shown.keys()].map((id, i) => [id, i]));
    for (const [id, t] of changes) {
      values.set(id, rich ? richValue(t) : t); shown.set(id, t);
      p.items[0].selectSingleNode(`/*[1]/*[${at.get(id) + 1}]`).text = values.get(id);
    }
  },
  // The common API (CustomXmlParts 1.1): CustomXMLNode.setTextAsync per node, one after the other
  // (calls issued together fail with "Wait until the previous call completes").
  async commonApi(c, changes) {
    const call = (o, m, ...a) => new Promise((res, rej) => o[m](...a, (r) => (r.status === "succeeded" ? res(r.value) : rej(r.error))));
    const p = await call(Office.context.document.customXmlParts, "getByIdAsync", ITEM_ID);
    const nodes = await call(p, "getNodesAsync", "/*/*");
    const at = new Map([...shown.keys()].map((id, i) => [id, i]));
    for (const [id, t] of changes) {
      values.set(id, rich ? richValue(t) : t); shown.set(id, t);
      await call(nodes[at.get(id)], "setTextAsync", values.get(id));
    }
  },
};

/** Office.js read of every control: how many show what they should. */
async function check(c) {
  const items = await bodyControls(c);
  let wrong = 0;
  const examples = [];
  for (const x of items) {
    if (!x.tag.startsWith(TAG)) continue;
    const want = shown.get(x.tag.slice(TAG.length));
    if (want !== undefined && x.text !== want) { wrong++; if (examples.length < 3) examples.push([x.tag, x.text, want]); }
  }
  return { wrong, examples };
}

/** Word's object model, the last three citations (far from the view): do they show the new text yet? */
async function comTail(n = 3) {
  const r = await command(`$cs = $doc.ContentControls; $o = @(); $k = 0; for ($i = $cs.Count; $i -ge 1 -and $k -lt ${n}; $i--) { $cc = $cs.Item($i); if ($cc.Tag -like '${TAG}*') { $o += ,@($cc.Tag, $cc.Range.Text); $k++ } }; $o`);
  const rows = Array.isArray(r?.[0]) ? r : [r];
  return rows.filter(Boolean).map(([tag, text]) => text === shown.get(tag.slice(TAG.length)));
}

/**
 * One write, timed: until the call returns; then an Office.js read of every
 * control right away; then a trivial call into Word's object model, which
 * waits while Word is busy; then Word's object model reads the last citations.
 */
async function timedWrite(name, writer, changes, extra = {}) {
  try {
    const r = await word(async (c) => {
      const t = T();
      await WRITERS[writer](c, changes);
      await c.sync();
      return { returnMs: ms(t), t };
    });
    const tj = T();
    const first = await word((c) => check(c));
    const jsCheckMs = ms(tj);
    const tp = T();
    await command("1");
    const comPingMs = ms(tp);
    const tail = await comTail();
    let verify = first;
    for (let k = 0; k < 20 && verify.wrong; k++) {
      await sleep(500);
      verify = await word((c) => check(c));
    }
    const row = { writer, changed: changes.size, returnMs: r.returnMs, jsCheckMs, jsWrongRightAfter: first.wrong, comPingMs, allShownMs: ms(r.t), comTailRight: tail, wrong: verify.wrong, examples: verify.examples, ...extra };
    log(name, row);
    return row;
  } catch (e) {
    log(name, { writer, changed: changes.size, error: err(e) });
    return false;
  }
}

// ---------- steps ----------
async function open() {
  log("requirement sets", { ...sets(), diagnostics: Office.context.diagnostics });
  await step("open: first read of every control (tag, text)", () => word(async (c) => ({ controls: (await bodyControls(c)).length })));
  await step("open: read the part", () => word(async (c) => { const p = part(c); await c.sync(); if (!p.items.length) return { parts: 0 }; const x = p.items[0].getXml(); await c.sync(); return { parts: p.items.length, id: p.items[0].id, bytes: x.value.length }; }));
  await step("open: check every control against the part", () => word((c) => check(c)));
}

async function timing() {
  const t = T();
  let round = 0;
  // Start from what the part holds.
  await word(async (c) => { const v = await partValues(c); for (const [id, x] of v) values.set(id, x); });
  // The cost of one Office.js text write into a bound control, and whether it writes back into the part.
  await step("insertText into 100 bound controls", () => word(async (c) => {
    const items = (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG)).slice(100, 200);
    const t0 = T();
    for (const x of items) x.insertText("[it]", "Replace");
    await c.sync();
    const perMs = +((T() - t0) / items.length).toFixed(1);
    const v = await partValues(c);
    const id = items[0].tag.slice(TAG.length);
    for (const x of items) { const i = x.tag.slice(TAG.length); shown.set(i, v.get(i) ?? "[it]"); values.set(i, rich ? richValue(shown.get(i)) : shown.get(i)); }
    return { perMs, partValueAfter: rich ? v.get(id)?.slice(-200) : v.get(id) };
  }));
  // Small changes first, with every writer; the whole document last, once per writer named in cfg.all.
  const counts = cfg.counts ?? ["one", 10, 100, 400];
  // A writer whose last write took more than a minute is not tried with more values.
  const slow = new Set();
  for (let run = 1; run <= cfg.runs; run++) {
    for (const f of counts) {
      for (const writer of cfg.writers ?? Object.keys(WRITERS)) {
        if (slow.has(writer)) { log(`write ${f} via ${writer}`, { skipped: "the last write with this writer took more than 60 s" }); continue; }
        const r = await timedWrite(`write ${f} via ${writer}`, writer, change(f, ++round), { run });
        if (!r || r.allShownMs > 60000) slow.add(writer);
      }
    }
    if (run === 1 && (cfg.writers ?? []).includes("setXml")) await timedWrite("write none via setXml", "setXml", change("none", ++round), { run });
    for (const writer of cfg.all ?? []) if (run === 1 || cfg.allEveryRun) await timedWrite(`write all via ${writer}`, writer, change("all", ++round), { run });
  }
  log("timing total", { ms: ms(t) });
}

async function undo() {
  await command("$doc.UndoClear(); 1");
  const before = new Map(shown);
  await timedWrite(`undo: write 400 via ${cfg.writer ?? "nodeSingle"}`, cfg.writer ?? "nodeSingle", change("tenth", "u"));
  for (let k = 1; k <= 4; k++) {
    const ok = await command("$doc.Undo(1)");
    const r = await word(async (c) => {
      const items = (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG));
      const v = await partValues(c);
      const reverted = items.filter((x) => x.text === before.get(x.tag.slice(TAG.length))).length;
      const partReverted = [...v].filter(([id, x]) => (rich ? x.includes(before.get(id)) : x === before.get(id))).length;
      return { reverted, partReverted, of: items.length, text0: items[0].text, part0: v.get(items[0].tag.slice(TAG.length))?.slice(-60) };
    });
    log(`undo: Ctrl+Z ${k}`, { undoReturned: ok, ...r });
    if (!ok) break;
  }
  // Put the part and the controls back in step with `shown`.
  const r = await word(async (c) => {
    const items = (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG));
    const v = await partValues(c);
    for (const x of items) { const id = x.tag.slice(TAG.length); values.set(id, v.get(id)); shown.set(id, rich ? x.text : v.get(id)); }
    return { controlsDifferFromPart: rich ? 0 : items.filter((x) => x.text !== v.get(x.tag.slice(TAG.length))).length };
  });
  log("undo: after the undos", r);
}

// The author is shown only when it is a neutral name; otherwise whether it is the Office user's name.
const revisions = () => command(`$o = @(); foreach ($r in $doc.Revisions) { $a = $r.Author; $o += ,(@($r.Type, $r.Range.Text, $(if ($a -like 'Author*' -or $a -eq '') { $a } elseif ($a -eq $word.UserName) { '(the Office user name)' } else { '(another name)' })) -join ' | ') }; @{ count = $doc.Revisions.Count; items = @($o | Select-Object -First 8) }`);

async function tracked() {
  await word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); });
  const changes = change("tenth", "t");
  const few = new Map([...changes].slice(0, 10));
  const before = new Map(shown);
  const W = cfg.writer ?? "nodeSingle";
  await timedWrite(`tracked: ${W} of 10 values with track changes on`, W, few);
  await timedWrite(`tracked: ${W} of the other 390 values with track changes on`, W, new Map([...changes].slice(10)));
  log("tracked: revisions (Word's object model: type, text, author is the Office user, author empty)", await revisions());
  await step("tracked: Office.js tracked changes", () => word(async (c) => {
    const t = c.document.body.getTrackedChanges();
    t.load("items/type,items/text,items/author");
    await c.sync();
    return { count: t.items.length, types: [...new Set(t.items.map((x) => x.type))] };
  }).catch((e) => ({ error: err(e) })));
  const [id1] = few.keys();
  await step("tracked: one changed control read three ways", () => word(async (c) => {
    const items = await bodyControls(c);
    const x = items.find((y) => y.tag === TAG + id1);
    const cur = x.getReviewedText("Current");
    const org = x.getReviewedText("Original");
    await c.sync();
    return { text: x.text, current: cur.value, original: org.value, before: before.get(id1) };
  }));
  for (const how of ["RejectAll", "AcceptAll"]) {
    if (how === "AcceptAll") await timedWrite(`tracked: ${W} of 10 values again`, W, new Map([...few].map(([id], k) => [id, `[ta.${k}]`])));
    const n = await command(`$doc.Revisions.${how}(); $doc.Revisions.Count`);
    await step(`tracked: after ${how}`, () => word(async (c) => {
      const items = await bodyControls(c, ",items/xmlMapping/isMapped");
      const v = await partValues(c);
      const rows = [...few.keys()].slice(0, 3).map((id) => { const x = items.find((y) => y.tag === TAG + id); return { text: x.text, part: rich ? v.get(id)?.slice(-80) : v.get(id), shownBefore: shown.get(id), mapped: x.xmlMapping.isMapped }; });
      const unmapped = items.filter((x) => x.tag.startsWith(TAG) && !x.xmlMapping.isMapped).length;
      return { revisionsLeft: n, rows, unmapped };
    }));
  }
  await word(async (c) => { c.document.changeTrackingMode = "Off"; await c.sync(); });
  // Does a bound update still go through after all that?
  await timedWrite(`tracked: ${W} of 10 values after accepting`, W, new Map([...few].map(([id], k) => [id, `[tz.${k}]`])));
  // The changed state is the part's.
  await word(async (c) => { const v = await partValues(c); for (const [id, x] of v) { values.set(id, x); if (!rich) shown.set(id, x); } });
}

async function reopen() {
  const s = cfg.stats;
  const ids = s.expectedFromStore ?? s.expectedFromDocument;
  await step("reopen: the three edited citations", () => word(async (c) => {
    const items = await bodyControls(c);
    const v = await partValues(c);
    return { rows: ids.map((id) => ({ id, control: items.find((x) => x.tag === TAG + id)?.text, part: v.get(id) })) };
  }));
  await step("reopen: controls that differ from the part", () => word(async (c) => {
    const items = (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG));
    const v = await partValues(c);
    return { differ: items.filter((x) => x.text !== v.get(x.tag.slice(TAG.length))).length, of: items.length };
  }));
}

// ---------- B: REF fields to numbered entries ----------
async function fields() {
  // The bibliographies as paragraphs, and the fields.
  await step("fields: count", () => word(async (c) => { const f = c.document.body.fields; f.load("items/code"); await c.sync(); return { fields: f.items.length, refs: f.items.filter((x) => /REF/.test(x.code)).length }; }));
  const cites = async (c) => (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG)).map((x) => x.text);
  const initial = await word(cites);
  log("fields: sample of citation texts", { first: initial.slice(0, 8), grouped: initial.filter((x) => x.includes(",")).slice(0, 4) });
  /** Rewrites each bibliography's entries in a new order, with their bookmarks, in one OOXML insert per bibliography. */
  const reorder = (name, how) => step(name, () => word(async (c) => {
    const bibs = (await bodyControls(c)).filter((x) => x.tag === BIB);
    const oo = bibs.map((b) => b.getRange("Content").getOoxml());
    await c.sync();
    let moved = 0;
    const t = T();
    bibs.forEach((b, i) => {
      const pkg = oo[i].value;
      const body = /<w:body>([\s\S]*?)(<w:sectPr[\s\S]*?<\/w:sectPr>)?<\/w:body>/.exec(pkg)[1];
      const ps = body.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
      const next = how(ps, i, bibs.length);
      moved += next.filter((p, k) => p !== ps[k]).length;
      b.getRange("Content").insertOoxml(pkg.replace(body, next.join("")), "Replace");
    });
    await c.sync();
    return { insertMs: ms(t), entriesMoved: moved, bibs: bibs.length };
  }));
  const update = async (name, ofWhich) => {
    const before = await word(cites);
    await step(name + ": Office.js updateResult on every REF field, one sync", () => word(async (c) => {
      const f = c.document.body.fields.getByTypes(["Ref"]);
      f.load("items");
      await c.sync();
      const t = T();
      for (const x of ofWhich ? ofWhich(f.items) : f.items) x.updateResult();
      await c.sync();
      const n = ofWhich ? ofWhich(f.items).length : f.items.length;
      return { fields: n, updateMs: ms(t), perFieldMs: +((T() - t) / n).toFixed(2) };
    }));
    const after = await word(cites);
    log(name + ": citations changed by the update", { changed: after.filter((x, i) => x !== before[i]).length, of: after.length, sample: after.slice(0, 5), errors: after.filter((x) => /Error|!/.test(x)).length });
  };
  if (cfg.quick) {
    // Word's own update of every field (not open to an add-in), then Office.js on 100 fields, one sync.
    await reorder("fields quick: reverse every bibliography", (ps) => [...ps].reverse());
    log("fields quick: before the update the reader sees", { sample: (await word(cites)).slice(0, 5) });
    const before = await word(cites);
    await step("fields quick: Word's object model, Fields.Update() on the body", () => command("$sw = [Diagnostics.Stopwatch]::StartNew(); [void]$doc.Fields.Update(); $sw.ElapsedMilliseconds"));
    const after = await word(cites);
    log("fields quick: citations changed by Fields.Update()", { changed: after.filter((x, i) => x !== before[i]).length, of: after.length, sample: after.slice(0, 5) });
    await reorder("fields quick: move one entry to the top", (ps) => { const k = Math.floor(ps.length / 10); return [ps[k], ...ps.slice(0, k), ...ps.slice(k + 1)]; });
    await update("fields quick: 100 fields", (items) => items.slice(0, 100));
  }
  for (let run = 1; run <= (cfg.quick ? 0 : cfg.runs); run++) {
    await reorder(`fields run ${run}: reverse every bibliography`, (ps) => [...ps].reverse());
    log(`fields run ${run}: before the update the reader sees`, { sample: (await word(cites)).slice(0, 5) });
    await update(`fields run ${run}, all`);
    // 10 %: the entry at 10 % of each bibliography moves to the top.
    await reorder(`fields run ${run}: move one entry to the top`, (ps) => { const k = Math.floor(ps.length / 10); return [ps[k], ...ps.slice(0, k), ...ps.slice(k + 1)]; });
    await update(`fields run ${run}, tenth`);
    // One: the last two entries of the last bibliography swap.
    await reorder(`fields run ${run}: swap the last two entries of the last bibliography`, (ps, i, n) => (i === n - 1 ? [...ps.slice(0, -2), ps.at(-1), ps.at(-2)] : ps));
    await update(`fields run ${run}, one`);
  }
  // Word's own update of every field, which an add-in cannot call: what Word itself costs.
  if (!cfg.quick) await step("fields: Word's object model, Fields.Update() on the whole body", () => command("$sw = [Diagnostics.Stopwatch]::StartNew(); [void]$doc.Fields.Update(); $sw.ElapsedMilliseconds"));
  // Only the fields that changed, one by one, through Office.js: the per-field cost for few fields.
  if (!cfg.quick) await update("fields: 100 fields, one by one", (items) => items.slice(0, 100));
  // Tracked changes on: a reorder and an update.
  await word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); });
  await reorder("fields tracked: swap the last two entries of the last bibliography", (ps, i, n) => (i === n - 1 ? [...ps.slice(0, -2), ps.at(-1), ps.at(-2)] : ps));
  log("fields tracked: revisions after the reorder", await revisions());
  await update("fields tracked: update all");
  log("fields tracked: revisions after the update", await revisions());
  await command("$doc.Revisions.RejectAll(); 1");
  await word(async (c) => { c.document.changeTrackingMode = "Off"; await c.sync(); });
  // A citation copied into another document (Word's FormattedText, no clipboard), then its fields updated there.
  log("fields: a citation copied into another document", await command(`$cc = $doc.ContentControls | Where-Object { $_.Tag -like '${TAG}*' } | Select-Object -First 1; $d2 = $word.Documents.Add(); $d2.Content.FormattedText = $cc.Range.FormattedText; $a = $d2.Content.Text; [void]$d2.Fields.Update(); $b = $d2.Content.Text; $d2.Close([ref]0); @($a, $b)`));
}

// ---------- C: insertText straight into the controls ----------
async function direct() {
  const cites = async (c) => (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG));
  // The direct way, for the table: insertText into one, 400 and all 4,000 rich-text controls, in one sync each.
  for (let run = 1; run <= cfg.runs; run++) {
    for (const [label, pick] of [["one", (a) => [a[Math.floor(a.length / 2)]]], ["tenth", (a) => a.filter((_, i) => i % 10 === 0)], ["all", (a) => a]]) {
      await step(`direct run ${run}: insertText into ${label}`, () => word(async (c) => {
        const a = pick(await cites(c));
        const t = T();
        a.forEach((x, i) => x.insertText(`[d${run}.${i}]`, "Replace"));
        await c.sync();
        return { n: a.length, syncMs: ms(t), perMs: +((T() - t) / a.length).toFixed(1) };
      }));
      log(`direct run ${run}: Word's object model ping after it`, await command("1"));
    }
  }
}

/** What a whole-body OOXML rewrite must keep: Word's object model counts and lists them. */
const snapshot = () => command(`
  $h = { param($s) $sha = [Security.Cryptography.SHA1]::Create(); [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($s))).Replace('-','').Substring(0,12) }
  $cc = $doc.ContentControls
  $tags = @(); $ids = @(); $texts = @(); foreach ($c in $cc) { $tags += $c.Tag; $ids += $c.ID; $texts += $c.Range.Text }
  $revs = @(); foreach ($r in $doc.Revisions) { $revs += ('' + $r.Type + ':' + $r.Range.Text + ':' + ($r.Author -eq $word.UserName) + ':' + $r.Date.ToString('s')) }
  $coms = @(); foreach ($c in $doc.Comments) { $coms += ($c.Range.Text + ':' + $c.Scope.Text.Length + ':' + $c.Date.ToString('s')) }
  $fns = @(); foreach ($f in $doc.Footnotes) { $fns += $f.Range.Text }
  $flds = @(); foreach ($f in $doc.Fields) { $flds += $f.Code.Text }
  $bms = @(); foreach ($b in $doc.Bookmarks) { $bms += ($b.Name + ':' + $b.Range.Text.Length) }
  @{ controls = $cc.Count; tags = (& $h ($tags -join '|')); ids = (& $h ($ids -join '|')); texts = (& $h ($texts -join '|')); revisions = $revs; comments = $coms; footnotes = $fns; fields = $flds; bookmarks = $bms; paragraphs = $doc.Paragraphs.Count; trackRevisions = $doc.TrackRevisions }`);

// ---------- C: one OOXML insert over the whole body ----------
async function bodyooxml() {
  // A tracked insertion, a tracked deletion, a comment, a bookmark, a footnote and a PAGE field, all in the body.
  log("bodyooxml: fixtures", await command(`
    [void]$doc.Bookmarks.Add('keepMe', $doc.Paragraphs(30).Range)
    [void]$doc.Comments.Add($doc.Paragraphs(40).Range.Words(3), 'Comment by Author B')
    $doc.TrackRevisions = $true
    $doc.Paragraphs(50).Range.InsertBefore('tracked words ')
    $doc.Paragraphs(60).Range.Words(2).Delete()
    $doc.TrackRevisions = $false
    $r = $doc.Paragraphs(70).Range; $r.Collapse(1); $fn = $doc.Footnotes.Add($r); $fn.Range.Text = 'A footnote about the survey.'
    $r = $doc.Paragraphs(80).Range; $r.Collapse(1); [void]$doc.Fields.Add($r, 33)
    1`));
  const before = await snapshot();
  log("bodyooxml: before", before);
  for (let run = 1; run <= cfg.runs; run++) {
    const label = `untracked, run ${run}`;
    await step(`bodyooxml: getOoxml of the body, every citation text changed, insertOoxml Replace (${label})`, () => word(async (c) => {
      const t0 = T();
      const oo = c.document.body.getOoxml();
      await c.sync();
      const readMs = ms(t0);
      let n = 0;
      const next = oo.value.replace(/(<w:tag w:val="exp:cite:[^"]*"\/>[\s\S]*?<w:sdtContent>[\s\S]*?<w:t(?: [^>]*)?>)\[[^<]*/g, (m, a) => { n++; return a + "[o" + run + "." + n + "]"; });
      const t = T();
      c.document.body.insertOoxml(next, "Replace");
      await c.sync();
      return { readMs, bytes: next.length, citationsChanged: n, insertMs: ms(t) };
    }));
    const after = await snapshot();
    const same = Object.fromEntries(Object.keys(before).map((k) => [k, JSON.stringify(before[k]) === JSON.stringify(after[k])]));
    log(`bodyooxml: after (${label})`, { same, after });
    const shownNow = await word(async (c) => (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG)).slice(0, 3).map((x) => x.text));
    log("bodyooxml: first citations now", shownNow);
  }
  // One Ctrl+Z after the last rewrite.
  log("bodyooxml: Ctrl+Z", await command("$doc.Undo(1)"));
  log("bodyooxml: first citations after Ctrl+Z", await word(async (c) => (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG)).slice(0, 3).map((x) => x.text)));
}

// ---------- C: other write paths on unbound controls ----------
async function paths() {
  const cites = async (c) => (await bodyControls(c)).filter((x) => x.tag.startsWith(TAG));
  await direct();
  const per = async (name, from, n, write) => step(name, () => word(async (c) => {
    const a = (await cites(c)).slice(from, from + n);
    const t = T();
    a.forEach((x, i) => write(x, i, c));
    await c.sync();
    return { n: a.length, perMs: +((T() - t) / a.length).toFixed(1) };
  }));
  const ooxmlRun = (text) => `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${esc(text)}</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  for (let run = 1; run <= cfg.runs; run++) {
    const o = run * 400;
    await per(`paths run ${run}: control.insertText Replace`, o, 100, (x) => x.insertText("[a]", "Replace"));
    await per(`paths run ${run}: content range insertText Replace`, o + 100, 100, (x) => x.getRange("Content").insertText("[b]", "Replace"));
    await per(`paths run ${run}: control.insertOoxml Replace`, o + 200, 100, (x) => x.insertOoxml(ooxmlRun("[c]"), "Replace"));
    await per(`paths run ${run}: first character range insertText (text shorter or same length)`, o + 300, 100, (x) => x.getRange("Content").search("[").getFirst().insertText("[", "Replace"));
  }
  await step("paths: Word's object model, 100 Range.Text", () => command(`$cs = $doc.ContentControls; $sw = [Diagnostics.Stopwatch]::StartNew(); for ($i = 2001; $i -le 2100; $i++) { $cs.Item($i).Range.Text = '[d]' }; $sw.ElapsedMilliseconds / 100`));
  // One OOXML insert over a range that holds many controls: a section's worth of paragraphs.
  // Before: a bookmark, a comment and a tracked insertion inside that range.
  await command(`$p = $doc.Paragraphs(30).Range; [void]$doc.Bookmarks.Add('keepMe', $p); [void]$doc.Comments.Add($doc.Paragraphs(40).Range, 'Comment by Author B'); $doc.TrackRevisions = $true; $doc.Paragraphs(50).Range.InsertBefore('tracked words '); $doc.TrackRevisions = $false; 1`);
  const marks = () => command(`@{ bookmarks = $doc.Bookmarks.Count; comments = $doc.Comments.Count; revisions = $doc.Revisions.Count; controls = $doc.ContentControls.Count }`);
  log("paths: before the range insert", await marks());
  for (const [label, track] of [["untracked", false], ["track changes on", true]]) {
    await step(`paths: one insertOoxml over paragraphs 20 to 400 (${label})`, () => word(async (c) => {
      if (track) { c.document.changeTrackingMode = "TrackAll"; await c.sync(); }
      const ps = c.document.body.paragraphs;
      ps.load("items");
      await c.sync();
      const range = ps.items[19].getRange("Start").expandTo(ps.items[399].getRange("End"));
      const ccs = range.contentControls;
      ccs.load("items");
      const oo = range.getOoxml();
      await c.sync();
      const before = ccs.items.length;
      // Every citation's text in the range changes, as in a renumbering.
      const next = oo.value.replace(/(<w:tag w:val="exp:cite:[^"]*"\/>[\s\S]*?<w:sdtContent>[\s\S]*?<w:t(?: [^>]*)?>)\[[^<]*/g, "$1[r]");
      const t = T();
      range.insertOoxml(next, "Replace");
      await c.sync();
      const insertMs = ms(t);
      if (track) { c.document.changeTrackingMode = "Off"; await c.sync(); }
      return { controlsInRange: before, bytes: next.length, insertMs, hasComment: /commentReference/.test(oo.value), hasIns: /<w:ins /.test(oo.value), hasBookmark: /w:name="keepMe"/.test(oo.value) };
    }));
    log(`paths: after the range insert (${label})`, await marks());
    await command("$doc.Undo(1); 1");
    log(`paths: after undoing it`, await marks());
  }
}

// ---------- limits: how large a part, how large one call ----------
async function limits() {
  for (const mb of [1, 5, 20, 50, 100]) {
    // mb × 1024 elements of 1,024 characters each.
    const xml = `<big xmlns="urn:siglum:exp:big">${("<c>" + "x".repeat(1017) + "</c>").repeat(mb * 1024)}</big>`;
    const added = await step(`limits: customXmlParts.add of ${mb} MB`, () => word(async (c) => {
      const p = c.document.customXmlParts.add(xml);
      p.load("id");
      await c.sync();
      return { id: p.id, chars: xml.length };
    }));
    if (!added?.id) break;
    await step(`limits: getXml of the ${mb} MB part`, () => word(async (c) => { const x = c.document.customXmlParts.getItem(added.id).getXml(); await c.sync(); return { chars: x.value.length }; }));
    await step(`limits: setXml of the ${mb} MB part`, () => word(async (c) => { c.document.customXmlParts.getItem(added.id).setXml(xml.replace("<c>x", "<c>y")); await c.sync(); return {}; }));
    await step(`limits: delete the ${mb} MB part`, () => word(async (c) => { c.document.customXmlParts.getItem(added.id).delete(); await c.sync(); return {}; }));
  }
}

// ---------- which controls Office.js sees ----------
async function mini() {
  log("mini: every story (Word's object model)", await command(`$o = @(); foreach ($cc in $doc.ContentControls) { $o += ,@($cc.Tag, $cc.Type, $cc.Range.Text, $cc.XMLMapping.IsMapped) }; $o`));
  for (const [name, get] of [["body.contentControls", (c) => c.document.body.contentControls], ["document.contentControls", (c) => c.document.contentControls], ["body.paragraphs first .contentControls", (c) => c.document.body.paragraphs.getFirst().contentControls], ["getContentControls()", (c) => c.document.body.getContentControls()], ["getByTag e1", (c) => c.document.contentControls.getByTag(TAG + "e1")], ["getByTag pu", (c) => c.document.contentControls.getByTag(TAG + "pu")]]) {
    await step("mini: " + name, () => word(async (c) => { const a = get(c); a.load("items/tag,items/text,items/type"); await c.sync(); return { rows: a.items.map((x) => [x.tag, x.type, x.text]) }; }));
  }
  await step("mini: getOoxml of the body: sdts in it", () => word(async (c) => { const x = c.document.body.getOoxml(); await c.sync(); return { sdts: (x.value.match(/<w:sdt>/g) ?? []).length, bindings: (x.value.match(/dataBinding/g) ?? []).length }; }));
}

// ---------- the small document: places, locks, formatting, copies, typing ----------
async function features() {
  // Every control in every story through Word's object model: story, tag, text, mapped, xpath.
  const stories = () => command(`$o = @(); foreach ($s in $doc.StoryRanges) { $r = $s; while ($r) { foreach ($cc in $r.ContentControls) { $o += ,@([int]$r.StoryType, $cc.Tag, $cc.Range.Text, $cc.XMLMapping.IsMapped) }; $r = $r.NextStoryRange } }; foreach ($c in $doc.Comments) { foreach ($cc in $c.Range.ContentControls) { $o += ,@('comment', $cc.Tag, $cc.Range.Text, $cc.XMLMapping.IsMapped) } }; foreach ($sh in $doc.Shapes) { if ($sh.TextFrame.HasText) { foreach ($cc in $sh.TextFrame.TextRange.ContentControls) { $o += ,@('shape', $cc.Tag, $cc.Range.Text, $cc.XMLMapping.IsMapped) } } }; $o`);
  const officeAll = () => word(async (c) => {
    const a = c.document.contentControls;
    a.load("items/tag,items/text");
    const b = c.document.body.getContentControls();
    b.load("items/tag,items/text,items/type");
    await c.sync();
    return { documentContentControls: a.items.map((x) => [x.tag, x.text]), bodyGetContentControls: b.items.map((x) => [x.tag, x.type, x.text]) };
  });
  const set = (pairs) => word(async (c) => {
    for (const [id, v] of pairs) values.set(id, v);
    const p = part(c);
    await c.sync();
    p.items[0].setXml(storeXml(values));
    await c.sync();
  });
  log("requirement sets", { ...sets(), diagnostics: Office.context.diagnostics });
  log("features: as opened, every story", await stories());
  log("features: as opened, Office.js document.contentControls", await officeAll());
  await step("features: xmlMapping through Office.js", () => word(async (c) => {
    const a = c.document.body.getContentControls();
    a.load("items/tag,items/xmlMapping/isMapped,items/xmlMapping/xpath,items/xmlMapping/prefixMappings");
    await c.sync();
    return { rows: a.items.map((x) => [x.tag, x.xmlMapping.isMapped, x.xmlMapping.xpath]) };
  }));
  // 7, 8: one write of the part; which places show it?
  await step("features: setXml of every plain value", () => set([...values.keys()].filter((id) => id !== "r1").map((id) => [id, `[new ${id}]`])));
  log("features: after setXml, every story", await stories());
  // 5: superscript and the rich control.
  log("features: superscript control after the update (Font.Superscript, -1 = yes)", await command(`$cc = $doc.SelectContentControlsByTag('${TAG}e4').Item(1); @($cc.Range.Text, $cc.Range.Font.Superscript)`));
  const etal = [{ text: "[new] Author B " }, { text: "et al.", italic: true }, { text: " 2021" }];
  await step("features: setXml of the rich control's package (italic et al.)", () => set([["r1", richPackage(etal.map((r) => runXml(r)).join(""))]]));
  const italics = () => command(`$cc = $doc.SelectContentControlsByTag('${TAG}r1').Item(1); $i = ''; foreach ($ch in $cc.Range.Characters) { if ($ch.Italic) { $i += $ch.Text } }; @($cc.Range.Text, $i, $cc.XMLMapping.IsMapped)`);
  log("features: rich control after the update (text, italic characters, mapped)", await italics());
  await step("features: insertOoxml into the rich control, then its element", () => word(async (c) => {
    const cc = (await bodyControls(c)).find((x) => x.tag === TAG + "r1");
    cc.insertOoxml(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Author C </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>et al.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`, "Replace");
    await c.sync();
    const v = await partValues(c);
    const r1 = v.get("r1");
    return { elementLength: r1.length, elementHasItalic: /<w:i\/>|<w:i /.test(r1), elementStart: r1.slice(0, 120), elementText: [...r1.matchAll(/<w:t[^>]*>([^<]*)</g)].map((m) => m[1]).join("") };
  }));
  // 4: typing into bound controls, plain and locked.
  const typeInto = (tag, text) => command(`$cc = $doc.SelectContentControlsByTag('${TAG}${tag}').Item(1); try { $cc.Range.Select(); $word.Selection.TypeText(${ps(text)}); 'typed' } catch { 'error: ' + $_.Exception.Message }`);
  const valuesNow = () => word(async (c) => Object.fromEntries(await partValues(c)));
  for (const tag of ["e1", "e2", "e3", "e11", "e5a"]) {
    const r = await typeInto(tag, `typed ${tag}`);
    const v = await valuesNow();
    log(`features: typing into ${tag}`, { result: r, element: v[tag.replace(/[abc]$/, "")], controls: (await stories()).filter((x) => /e1$|e1c|e2$|e3$|e11|e5/.test(x[1])).map((x) => [x[1], x[2]]) });
  }
  await step("features: Office.js insertText into the content-locked e2", () => word(async (c) => { (await bodyControls(c)).find((x) => x.tag === TAG + "e2").insertText("[js]", "Replace"); await c.sync(); return "written"; }));
  await step("features: setXml with e2, e3, e11 changed (locked controls)", () => set([["e2", "[part e2]"], ["e3", "[part e3]"], ["e11", "[part e11]"], ["e5", "[part e5]"]]));
  log("features: locked and double after setXml", (await stories()).filter((x) => /e2$|e3$|e11|e5/.test(x[1])).map((x) => [x[1], x[2]]));
  // 6: copies through Word's FormattedText (Word's own copy, without the clipboard).
  const range = (tag) => `$cc = $doc.SelectContentControlsByTag('${TAG}${tag}').Item(1); $src = $doc.Range($cc.Range.Start - 1, $cc.Range.End + 1)`;
  log("features: copy e1 to the end of the same document", await command(`${range("e1")}; $end = $doc.Content; $end.InsertParagraphAfter(); $dst = $doc.Paragraphs.Last.Range; $dst.FormattedText = $src.FormattedText; $n = $doc.SelectContentControlsByTag('${TAG}e1').Count; $c2 = $doc.SelectContentControlsByTag('${TAG}e1').Item($n); @($n, $c2.Range.Text, $c2.XMLMapping.IsMapped, $c2.XMLMapping.XPath)`));
  log("features: copy e1 into a new document", await command(`${range("e1")}; $d2 = $word.Documents.Add(); $d2.Content.FormattedText = $src.FormattedText; $c2 = $d2.ContentControls.Item(1); $o = @($d2.ContentControls.Count, $c2.Range.Text, $c2.XMLMapping.IsMapped, $c2.XMLMapping.XPath, $d2.CustomXMLParts.Count, @($d2.CustomXMLParts | ForEach-Object { $_.NamespaceURI })); $d2.Close([ref]0); $o`));
  log("features: move e4 (cut and paste: copy, then delete the original)", await command(`${range("e4")}; $dst = $doc.Paragraphs.Last.Range; $dst.InsertParagraphAfter(); $dst = $doc.Paragraphs.Last.Range; $dst.FormattedText = $src.FormattedText; $src.Delete(); $c2 = $doc.SelectContentControlsByTag('${TAG}e4').Item(1); @($doc.SelectContentControlsByTag('${TAG}e4').Count, $c2.Range.Text, $c2.XMLMapping.IsMapped)`));
  log("features: move the paragraph with e6? no: the first paragraph (e1) down with track changes on", await command(`$doc.TrackRevisions = $true; $doc.Paragraphs(1).Range.Relocate(1); $doc.TrackRevisions = $false; $o = @(); foreach ($cc in $doc.SelectContentControlsByTag('${TAG}e1')) { $o += ,@($cc.Range.Text, $cc.XMLMapping.IsMapped) }; @{ revisions = $doc.Revisions.Count; e1 = $o }`));
  await step("features: setXml after the copies and moves", () => set([["e1", "[after e1]"], ["e4", "[after e4]"]]));
  log("features: copies after setXml", (await stories()).filter((x) => /e1$|e4$/.test(x[1])).map((x) => [x[0], x[1], x[2], x[3]]));
  // 3 on the small document: who is the revision's author?
  await word(async (c) => { c.document.changeTrackingMode = "TrackAll"; await c.sync(); });
  await step("features: setXml with track changes on", () => set([["e6", "[tracked e6]"], ["e7", "[tracked e7]"]]));
  log("features: revisions", await revisions());
  await step("features: Office.js tracked changes after the bound update", () => word(async (c) => { const t = c.document.body.getTrackedChanges(); t.load("items/type,items/text"); await c.sync(); return { rows: t.items.map((x) => [x.type, x.text]) }; }));
  await word(async (c) => { c.document.changeTrackingMode = "Off"; await c.sync(); });
  // A tracked typed change in a bound control: what reaches the part, and what accepting or rejecting does.
  await command("$doc.TrackRevisions = $true; 1");
  log("features: typing into e6 with track changes on", await typeInto("e6", "typed tracked e6"));
  await command("$doc.TrackRevisions = $false; 1");
  log("features: revisions after the tracked typing", await revisions());
  log("features: element e6 after tracked typing", (await valuesNow()).e6);
  await command("$doc.Revisions.RejectAll(); 1");
  log("features: after RejectAll: e6 control and element", { control: (await stories()).filter((x) => /e6$/.test(x[1])).map((x) => x[2]), element: (await valuesNow()).e6 });
  // 9: undo after a bound update.
  await command("$doc.UndoClear(); 1");
  await set([["e8", "[undo e8]"], ["e9", "[undo e9]"]]);
  for (let k = 1; k <= 3; k++) {
    const ok = await command("$doc.Undo(1)");
    const v = await valuesNow();
    log(`features: Ctrl+Z ${k} after a setXml`, { undoReturned: ok, e8: v.e8, e9: v.e9, controls: (await stories()).filter((x) => /e8$|e9$/.test(x[1])).map((x) => [x[1], x[2]]) });
    if (!ok) break;
  }
}

onReady(async () => {
  cfg = await (await fetch("bind-config.json", { cache: "no-store" })).json();
  rich = cfg.stats?.variant === "rich";
  values = new Map(cfg.cites.map(([id, text]) => [id, rich ? richValue(text) : text]));
  shown = new Map(cfg.cites.map(([id, text]) => [id, text]));
  if (cfg.stats?.variant === "features") { values = new Map(cfg.cites); shown = new Map(); }
  const steps = { direct, bodyooxml, limits, mini, open, timing, undo, tracked, reopen, fields, paths, features };
  for (const s of cfg.steps) {
    try { await steps[s](); } catch (e) { log(`step ${s} failed`, err(e) + "\n" + e?.stack); }
  }
});
