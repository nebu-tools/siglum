// A small stand-in for the parts of Word's API that the legacy upgrade uses.
// Like Word, it applies changes at sync() and gives values only after it.

/** A document with custom XML parts, content controls and paragraphs of text. */
export function fakeDocument({ parts = [], controls = [], paragraphs = [] } = {}) {
  const doc = { parts: parts.map((xml) => ({ xml })), controls: controls.map((c) => ({ ...c })), paragraphs: [...paragraphs], syncs: 0, tracking: "TrackAll", log: [] };
  const pending = [];
  const later = (fn) => pending.push(fn);
  const namespace = (xml) => /xmlns="([^"]+)"/.exec(xml)?.[1];

  const partProxy = (part) => ({
    getXml() { const r = {}; later(() => { r.value = part.xml; }); return r; },
    setXml(xml) { later(() => { part.xml = xml; }); },
    delete() { later(() => { doc.parts = doc.parts.filter((p) => p !== part); }); },
  });
  const collection = (fill) => {
    const c = { items: undefined, load() { later(() => { c.items = fill(); }); } };
    return c;
  };
  const controlProxy = (control) => ({
    get tag() { return control.tag; },
    set tag(v) { later(() => { control.tag = v; }); },
    get title() { return control.title; },
    set title(v) { later(() => { control.title = v; }); },
  });
  // Word's wildcards, as far as the upgrade uses them: "\x" is a literal x, "*" is any text.
  const pattern = (text, wild) => new RegExp(wild
    ? text.replace(/\\(.)|\*|[.+?^${}()|[\]]/g, (m, lit) => (lit ? "\\" + lit : m === "*" ? ".*?" : "\\" + m))
    : text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
  const search = (text, { matchWildcards = false } = {}) => collection(() => doc.paragraphs.flatMap((p, i) => [...p.matchAll(pattern(text, matchWildcards))].map((m) => {
    const range = {
      text: m[0],
      insertText(next, where) {
        if (where !== "Replace") throw new Error("the fake replaces only");
        doc.log.push({ replace: m[0], with: next, tracking: doc.tracking });
        range.edit = { paragraph: i, start: m.index, end: m.index + m[0].length, next };
        later(() => {});
      },
    };
    edits.push(range);
    return range;
  })));
  const edits = [];

  const context = {
    document: {
      customXmlParts: {
        getByNamespace: (ns) => collection(() => doc.parts.filter((p) => namespace(p.xml) === ns).map(partProxy)),
        add(xml) { later(() => { doc.parts.push({ xml }); }); },
      },
      body: {
        get contentControls() { return collection(() => doc.controls.map(controlProxy)); },
        search,
      },
    },
    async sync() {
      doc.syncs++;
      for (const fn of pending.splice(0)) fn();
      // Later ranges first, so the earlier positions stay valid.
      const todo = edits.filter((r) => r.edit).map((r) => r.edit).sort((a, b) => b.paragraph - a.paragraph || b.start - a.start);
      for (const e of todo) doc.paragraphs[e.paragraph] = doc.paragraphs[e.paragraph].slice(0, e.start) + e.next + doc.paragraphs[e.paragraph].slice(e.end);
      for (const r of edits) delete r.edit;
    },
  };
  return { doc, context };
}
