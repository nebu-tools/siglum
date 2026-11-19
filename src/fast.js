// The fast whole-body path of refresh (see word.js): many citations rewritten
// with one body OOXML write instead of one write per control, verified
// against the plan, with the per-control path as fallback. Works without
// Word; the Office.js calls run against the given context.

import { CITE_PREFIX, BIB_TAG } from "./model.js";
import { bookmarkNames, rewriteCitationTexts, runsText } from "./ooxml.js";

const isCite = (cc) => cc.tag?.startsWith(CITE_PREFIX);
const isOurs = (cc) => isCite(cc) || cc.tag === BIB_TAG;

/** Citations rewritten by one whole-body OOXML write instead of one write per
 * control when at least this many need it. Rewriting the whole body changed
 * 4,000 citations in 5.5–6.9 s against 106–341 s one by one, with track
 * changes off: the fixed cost is about one OOXML read, the write and one
 * verification read. Below this count the per-control path is no slower.
 */
export const FAST_CITES_MIN = 500;

// Word's own bookmarks start with an underscore; a person's own bookmarks do
// not. A whole-body write has grown a person's bookmark by a character per
// write, so the fast path stays off while the document holds any of theirs.
// Images and other embedded objects pass through byte-identical, but this is
// untested with them, with threaded comments and with co-authoring.
const personBookmark = (name) => !name.startsWith("_");

/**
 * Whether the fast whole-body path may run: no tracked changes, enough
 * citations to write, and, unless `withOthers` says their writes already
 * went through the existing path, nothing but citations. Bibliographies keep
 * their own one-piece path, and leftovers their deletes.
 */
export function fastEligible({ tracked, bibs, removes, cites, withOthers = false }) {
  return !tracked && (withOthers || (!bibs && !removes)) && cites >= FAST_CITES_MIN;
}

/**
 * Rewrites the citations in `rest` with one body OOXML write, then checks
 * every rewritten control's text against what the plan wants. `planned`
 * are all the plan's citations with their positions, `restCites` those
 * still to write, and `rest` maps them to content control ids. Returns
 * the citations that still need the per-control path: all of them when the
 * fast path does not apply or its write fails, the mismatches when
 * verification finds any. Never throws: the per-control path below writes
 * what is left.
 */
export async function fastCites(context, ccs, controls, planned, restCites, rest) {
  try {
    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();
    if (doc.changeTrackingMode !== "Off") return rest;
    const oo = doc.body.getOoxml();
    await context.sync();
    if (bookmarkNames(oo.value).some(personBookmark)) return rest;
    const rewrites = new Map(restCites.map((c) => [ccs[c.index].tag, c.runs]));
    const { xml, applied, skipped } = rewriteCitationTexts(oo.value, rewrites);
    if (skipped.length || applied.length !== restCites.length) return rest;
    doc.body.insertOoxml(xml, "Replace");
    await context.sync();
    // Every citation the plan writes, including the fresh ones already
    // written above, reads back as planned; the rest stays as it was.
    const wantRuns = new Map(planned.map((c) => [c.index, runsText(c.runs)]));
    const want = new Map(ccs.map((cc, i) => [cc.id, wantRuns.has(i) ? wantRuns.get(i) : controls[i].current]));
    const check = doc.body.contentControls;
    check.load("items/id,items/tag,items/text");
    await context.sync();
    const missing = new Set(check.items
      .filter((cc) => isOurs(cc) && want.has(cc.id) && want.get(cc.id) !== cc.text)
      .map((cc) => cc.id));
    return rest.filter((c) => missing.has(c.id));
  } catch (e) {
    console.error(e);
    return rest;
  }
}
