// Reads what Siglum needs out of OOXML, as Word's getOoxml returns it: the
// texts of content controls and a paragraph's style. Works without Word.
//
// The XML goes through a strict, namespace-aware parser. Elements are known
// by namespace and local name, so any prefix bound to Word's namespace reads
// the same as "w:". Input that is not well-formed, or that ends before its
// root element does, throws an OoxmlError, and nothing is read from it.

import { SaxesParser } from "saxes";
import { UserError } from "./messages.js";

/** WordprocessingML's main namespace. */
export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// The package wrapper that getOoxml returns, and markup compatibility.
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAIN_PART = "/word/document.xml";

/**
 * Siglum could not read the document's OOXML, or what it read does not match
 * what Word lists. The operation stops before it writes anything. `cause`
 * holds the parser's message.
 */
export class OoxmlError extends UserError {
  constructor(key, params = {}, cause = undefined) {
    super(key, params);
    this.name = "OoxmlError";
    if (cause !== undefined) this.cause = cause;
  }
}

class ParseFailure extends Error {}

/**
 * Parses `xml` and calls `open(node, parent)`, `close(node)` and `text(s)` for
 * the main document part: the content of the package's "/word/document.xml"
 * part, or all of `xml` when its root is a WordprocessingML element. The
 * package's other parts are parsed and checked, but not read; their binary
 * data can be large. A node is saxes's {uri, local, attributes}.
 */
function read(xml, { open = () => {}, close = () => {}, text = () => {} }) {
  const parser = new SaxesParser({ xmlns: true, position: false });
  parser.on("error", (e) => { throw new ParseFailure(e.message); });
  let depth = 0;
  let from = -1; // the depth below which the main part's elements are; -1 when outside it
  let main = false; // inside the main part's pkg:part
  let found = false;
  const stack = [];
  parser.on("opentag", (node) => {
    depth++;
    if (from >= 0) {
      open(node, stack[stack.length - 1]);
      stack.push(node);
    } else if (depth === 1) {
      if (node.uri === W) { from = 0; found = true; open(node, undefined); stack.push(node); }
      else if (node.uri !== PKG || node.local !== "package") throw new ParseFailure(`unexpected root element ${node.name}`);
    } else if (depth === 2 && node.uri === PKG && node.local === "part") {
      main = attr(node, PKG, "name") === MAIN_PART;
    } else if (depth === 3 && main && node.uri === PKG && node.local === "xmlData") {
      if (found) throw new ParseFailure(`two parts named ${MAIN_PART}`);
      from = depth;
      found = true;
    }
  });
  parser.on("closetag", (node) => {
    if (from >= 0 && depth > from) { stack.pop(); close(node); }
    else if (depth === from) from = -1;
    else if (depth === 2) main = false;
    depth--;
  });
  const onText = (s) => { if (from >= 0 && depth > from) text(s); };
  parser.on("text", onText);
  parser.on("cdata", onText);
  try {
    parser.write(xml).close();
  } catch (e) {
    if (e instanceof ParseFailure) throw new OoxmlError("error.ooxml", {}, e.message);
    throw e;
  }
  if (!found) throw new OoxmlError("error.ooxml", {}, `no part named ${MAIN_PART}`);
}

/** The value of the attribute `local` in namespace `uri`. */
function attr(node, uri, local) {
  for (const a of Object.values(node.attributes)) if (a.uri === uri && a.local === local) return a.value;
  return undefined;
}

// An on/off property such as <w:showingPlcHdr/> is on unless its w:val says off.
const on = (node) => !["0", "false", "off"].includes(attr(node, W, "val") ?? "1");

// Revision marks: their content is only in the text with changes accepted
// ("current"), or only in the text with changes rejected ("original").
const ONLY_CURRENT = new Set(["ins", "moveTo"]);
const ONLY_ORIGINAL = new Set(["del", "moveFrom"]);
// What an open element undoes when it closes.
const NONE = 0, SDT = 1, SDT_PR = 2, CURRENT = 3, ORIGINAL = 4, PROPS = 5, P = 6, R = 7, TEXT = 8, SKIP = 9;

/**
 * The content controls in the main document part of `ooxml`, in document
 * order, each as {tag, current, original}: its text with tracked changes
 * accepted and with them rejected, "" for both while it shows its
 * placeholder. Paragraphs end in "\r", a tab is "\t" and a line break "\v",
 * as in Word's `text`; a paragraph mark that is a tracked insertion or
 * deletion counts only on its side. Controls in text boxes are left out,
 * because Word's body.contentControls leaves them out too, and so is the
 * fallback of markup-compatibility content, which repeats its first choice.
 *
 * Unlike a control's `text`, this does not depend on how Word shows tracked
 * changes, and unlike getReviewedText it tells a move from a deletion.
 * Throws an OoxmlError when `ooxml` is not well-formed.
 */
export function controlTexts(ooxml) {
  const out = [];
  const controls = []; // the controls around the reading position
  const undo = []; // per open element, what its closing undoes
  const paragraphs = []; // per open paragraph, on which side its mark counts: {current, original}
  let current = 0; // nesting of revision marks that only the current text has
  let original = 0;
  let props = 0; // nesting of property elements, whose revision marks are about formatting
  let run = 0; // inside a run, where a tab or a break is text
  let text = 0; // inside w:t or w:delText
  let skip = 0; // nesting inside a text box or a fallback, which is not read
  let sdtPr = null; // the control whose properties are being read
  const add = (s, toCurrent = !original, toOriginal = !current) => {
    for (const c of controls) {
      if (toCurrent) c.current += s;
      if (toOriginal) c.original += s;
    }
  };
  const kind = (node, parent) => {
    if (node.uri === MC) return node.local === "Fallback" ? SKIP : NONE;
    if (node.uri !== W) return NONE;
    const name = node.local;
    if (name === "txbxContent") return SKIP;
    if (name === "sdt") return SDT;
    if (name === "sdtPr") return SDT_PR;
    if (sdtPr && parent?.local === "sdtPr" && parent.uri === W) {
      if (name === "tag") sdtPr.tag = attr(node, W, "val") ?? "";
      else if (name === "showingPlcHdr") sdtPr.placeholder = on(node);
    }
    if (ONLY_CURRENT.has(name) || ONLY_ORIGINAL.has(name)) {
      // In a paragraph's run properties it marks the paragraph mark itself.
      if (parent?.local === "rPr" && parent.uri === W && paragraphs.length && props === 2) {
        const mark = paragraphs[paragraphs.length - 1];
        if (ONLY_CURRENT.has(name)) mark.original = false;
        else mark.current = false;
        return NONE;
      }
      if (props) return NONE;
      return ONLY_CURRENT.has(name) ? CURRENT : ORIGINAL;
    }
    if (name.endsWith("Pr")) return PROPS;
    if (name === "p") return P;
    if (name === "r") return R;
    if (name === "t" || name === "delText") return TEXT;
    if (run && !text && !props) {
      if (name === "tab") add("\t");
      else if (name === "br" || name === "cr") add("\v");
    }
    return NONE;
  };
  read(ooxml, {
    open(node, parent) {
      if (skip) { skip++; undo.push(NONE); return; }
      const k = kind(node, parent);
      undo.push(k);
      if (k === SKIP) skip = 1;
      else if (k === SDT) {
        const entry = { tag: null, current: "", original: "", placeholder: false };
        out.push(entry);
        controls.push(entry);
      } else if (k === SDT_PR) { sdtPr = controls[controls.length - 1]; props++; }
      else if (k === CURRENT) current++;
      else if (k === ORIGINAL) original++;
      else if (k === PROPS) props++;
      else if (k === P) paragraphs.push({ current: true, original: true });
      else if (k === R) run++;
      else if (k === TEXT) text++;
    },
    close() {
      const k = undo.pop();
      if (skip) { skip--; return; }
      if (k === SDT) {
        const entry = controls.pop();
        if (entry.placeholder) { entry.current = ""; entry.original = ""; }
        delete entry.placeholder;
      } else if (k === SDT_PR) { sdtPr = null; props--; }
      else if (k === CURRENT) current--;
      else if (k === ORIGINAL) original--;
      else if (k === PROPS) props--;
      else if (k === P) {
        const mark = paragraphs.pop();
        add("\r", mark.current && !original, mark.original && !current);
      } else if (k === R) run--;
      else if (k === TEXT) text--;
    },
    text(s) {
      if (text && !skip) add(s);
    },
  });
  return out;
}

/**
 * The style id of the first paragraph in `ooxml`, from its w:pStyle, or null.
 * Throws an OoxmlError when `ooxml` is not well-formed.
 */
export function paragraphStyle(ooxml) {
  let style = null;
  let seen = false; // the first paragraph has begun
  let depth = 0; // nesting inside the first paragraph: 1 for the paragraph itself
  let pPr = false; // its own properties are open
  read(ooxml, {
    open(node) {
      if (depth) depth++;
      else if (!seen && node.uri === W && node.local === "p") { seen = true; depth = 1; return; }
      const isW = node.uri === W;
      if (depth === 2) pPr = isW && node.local === "pPr";
      else if (depth === 3 && pPr && isW && node.local === "pStyle") style = attr(node, W, "val") ?? null;
    },
    close() {
      if (depth) depth--;
      if (depth < 2) pPr = false;
    },
  });
  return style;
}
