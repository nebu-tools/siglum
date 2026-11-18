// Reads content controls' texts out of a body's OOXML, as Word's file format
// holds them. Works without Word.

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" };
const decode = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (all, e) => {
  if (e[0] === "#") return String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : +e.slice(1));
  return ENTITIES[e] ?? all;
});

// Revision marks: their content is only in the text with changes accepted
// ("current"), or only in the text with changes rejected ("original").
const ONLY_CURRENT = new Set(["w:ins", "w:moveTo"]);
const ONLY_ORIGINAL = new Set(["w:del", "w:moveFrom"]);

/**
 * The content controls in `ooxml` in document order, each as {tag, current,
 * original}: its text with tracked changes accepted and with them rejected,
 * "" for both while it shows its placeholder. Paragraphs end in "\r" and a
 * tab is "\t", as in Word's `text`. Controls in text boxes are left out,
 * because Word's body.contentControls leaves them out too.
 *
 * Unlike a control's `text`, this does not depend on how Word shows tracked
 * changes, and unlike getReviewedText it tells a move from a deletion.
 */
export function controlTexts(ooxml) {
  const out = [];
  const open = []; // the controls around the reading position: {entry}
  let current = 0; // nesting of revision marks that only the current text has
  let original = 0;
  let box = 0; // nesting of text boxes
  let field = 0; // inside field instructions
  let run = 0; // inside a run, where a tab or a break is text
  let inText = null; // the element whose text is being read
  let sdtPr = null; // the control whose properties are being read
  const add = (s) => {
    if (box) return;
    for (const c of open) {
      if (!original) c.entry.current += s;
      if (!current) c.entry.original += s;
    }
  };
  const re = /<(\/?)([\w:]+)([^>]*?)(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(ooxml))) {
    if (m[5] != null) {
      if (inText && !field) add(decode(m[5]));
      continue;
    }
    const [, close, name, attrs, empty] = m;
    if (close) {
      if (name === "w:t" || name === "w:delText") inText = null;
      else if (name === "w:instrText") field--;
      else if (ONLY_CURRENT.has(name)) current--;
      else if (ONLY_ORIGINAL.has(name)) original--;
      else if (name === "w:txbxContent") box--;
      else if (name === "w:sdtPr") sdtPr = null;
      else if (name === "w:p") add("\r");
      else if (name === "w:r") run--;
      else if (name === "w:sdt") {
        const c = open.pop();
        if (c?.entry.placeholder) { c.entry.current = ""; c.entry.original = ""; }
        if (c) delete c.entry.placeholder;
      }
      continue;
    }
    if (name === "w:sdt") {
      if (empty) continue;
      const entry = { tag: null, current: "", original: "", placeholder: false };
      if (!box) { out.push(entry); open.push({ entry }); } else open.push({ entry: { current: "", original: "" } });
      sdtPr = entry;
      continue;
    }
    if (sdtPr && name === "w:tag") sdtPr.tag = decode(/w:val="([^"]*)"/.exec(attrs)?.[1] ?? "");
    else if (sdtPr && name === "w:showingPlcHdr") sdtPr.placeholder = true;
    if (empty) {
      if (name === "w:tab" && run) add("\t");
      else if ((name === "w:br" || name === "w:cr") && run) add("\v");
      else if (name === "w:p") add("\r");
      continue;
    }
    if (name === "w:r") run++;
    if (name === "w:t" || name === "w:delText") inText = name;
    else if (name === "w:instrText") field++;
    else if (ONLY_CURRENT.has(name)) current++;
    else if (ONLY_ORIGINAL.has(name)) original++;
    else if (name === "w:txbxContent") box++;
  }
  return out;
}
