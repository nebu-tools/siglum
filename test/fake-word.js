// A small stand-in for the parts of Word's API that locating typed markers
// uses. Like Word, it queues work and gives values only after sync().
//
// The document is a list of paragraphs. A position is an offset into their
// texts joined by "\r", as Word separates paragraphs. A content control is a
// span {tag, start, end}; a control inside another one's span is nested in it.
//
// Its literal search is Word's as far as this code relies on it: matches in
// document order and without overlaps, case-sensitive, at most 255
// characters, and "^^" for a caret. Any other "^" code is refused, so a test
// fails where real Word would read a special code.

const SEARCH_MAX = 255;

export function fakeDocument({ paragraphs = [], controls = [] } = {}) {
  const doc = { paragraphs: [...paragraphs], controls: controls.map((c) => ({ ...c })), syncs: 0, searches: [] };
  const text = doc.paragraphs.join("\r");
  const starts = [];
  doc.paragraphs.reduce((at, p) => { starts.push(at); return at + p.length + 1; }, 0);
  let pending = [];
  const later = (fn) => pending.push(fn);

  const contains = (outer, inner) => outer.start <= inner.start && inner.end <= outer.end;
  /** The innermost control around a span, other than the span's own control. */
  const parentOf = (span, self) => doc.controls
    .filter((c) => c !== self && contains(c, span))
    .sort((a, b) => (a.end - a.start) - (b.end - b.start))[0];

  const nullable = (control) => {
    const proxy = control ? controlProxy(control) : {};
    proxy.isNullObject = !control;
    proxy.load = () => {};
    return proxy;
  };
  function controlProxy(control) {
    return {
      tag: control.tag,
      text: text.slice(control.start, control.end),
      getRange() { return rangeProxy(control.start, control.end); },
      get parentContentControlOrNullObject() { return nullable(parentOf(control, control)); },
    };
  }
  function rangeProxy(start, end) {
    const r = {
      start, end,
      load(props) {
        if (props !== "text") throw new Error(`the fake loads only text, not ${props}`);
        later(() => { r.text = text.slice(start, end); });
      },
      expandTo(other) { return rangeProxy(Math.min(start, other.start), Math.max(end, other.end)); },
      compareLocationWith(other) {
        const res = {};
        later(() => { res.value = relation(r, other); });
        return res;
      },
      get parentContentControlOrNullObject() { return nullable(parentOf(r)); },
      get paragraphs() { return { getFirst: () => paragraphProxy(paragraphAt(start)) }; },
    };
    return r;
  }
  const collection = (fill) => {
    const c = { items: undefined, load() { later(() => { c.items = fill(); }); } };
    return c;
  };

  /** Word's literal search for `query` in the paragraphs `from`…`to`. */
  const search = (query, options, from, to) => collection(() => {
    doc.searches.push(query);
    if (!options?.matchCase) throw new Error("the fake searches with matchCase only");
    if (query.length > SEARCH_MAX) throw new Error("SearchStringInvalidOrTooLong");
    const needle = query.replace(/\^(.?)/g, (all, c) => {
      if (c === "^") return "^";
      throw new Error(`the fake knows no special code ${all}`);
    });
    const hits = [];
    for (let p = from; p <= to; p++) {
      const t = doc.paragraphs[p];
      for (let i = t.indexOf(needle); i !== -1 && needle; i = t.indexOf(needle, i + needle.length)) {
        hits.push(rangeProxy(starts[p] + i, starts[p] + i + needle.length));
      }
    }
    return hits;
  });

  const paragraphProxy = (p) => {
    const q = rangeProxy(starts[p], starts[p] + doc.paragraphs[p].length);
    q.load = (props) => {
      if (props !== "text") throw new Error(`the fake loads only text, not ${props}`);
      later(() => { q.text = doc.paragraphs[p]; });
    };
    q.search = (query, options) => search(query, options, p, p);
    // Like Word, a paragraph compares its place only through a range.
    const compare = q.compareLocationWith;
    delete q.compareLocationWith;
    q.getRange = () => ({ ...q, compareLocationWith: compare });
    return q;
  };
  /** The index of the paragraph that holds position `at`. */
  const paragraphAt = (at) => starts.findLastIndex((x) => x <= at);

  const context = {
    document: {
      body: {
        search: (query, options) => search(query, options, 0, doc.paragraphs.length - 1),
        get paragraphs() { return collection(() => doc.paragraphs.map((_, p) => { const q = paragraphProxy(p); q.text = doc.paragraphs[p]; return q; })); },
        get contentControls() { return collection(() => [...doc.controls].sort((a, b) => a.start - b.start).map(controlProxy)); },
      },
    },
    async sync() {
      doc.syncs++;
      const todo = pending;
      pending = [];
      for (const fn of todo) fn();
    },
  };
  return { doc, context };
}

/** Word's LocationRelation of range `a` to range `b`. */
function relation(a, b) {
  if (a.start === b.start && a.end === b.end) return "Equal";
  if (b.start <= a.start && a.end <= b.end) return a.start === b.start ? "InsideStart" : a.end === b.end ? "InsideEnd" : "Inside";
  if (a.start <= b.start && b.end <= a.end) return a.start === b.start ? "ContainsStart" : a.end === b.end ? "ContainsEnd" : "Contains";
  if (a.end === b.start) return "AdjacentBefore";
  if (b.end === a.start) return "AdjacentAfter";
  if (a.end < b.start) return "Before";
  if (b.end < a.start) return "After";
  return a.start < b.start ? "OverlapsBefore" : "OverlapsAfter";
}
