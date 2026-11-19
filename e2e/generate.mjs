// Generates a large test document: pages of text in B sections, C citations
// spread through them, a bibliography at the end of each section and R
// sources, all in the add-in's stored form (names from src/model.js). The
// numbers and bibliographies come from the add-in's own plan() and citeproc,
// so the document is current when Word opens it. With --stale every citation
// reads "[?]" and every bibliography "…".
//
//   node e2e/generate.mjs --pages 20 --citations 100 --bibs 1 --refs 50 [--stale] [--seed 1] [--pane] out.docx
//
// --pane adds a task pane for the measuring page that opens with the document.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRenderer } from "../src/render.js";
import { mergeItems } from "../src/library.js";
import { CITE_PREFIX, BIB_TAG, CITE_NEW, emptyStore, storeToXml, citeTitle, newId, plan } from "../src/model.js";
import { documentText } from "../src/messages.js";
import { root, XML, W, EMPTY, esc, addPane, writeDocx, idFor } from "./rig.mjs";

/** The add-in's renderer with the style and locales that taskpane.js imports. */
export function nodeRenderer() {
  const src = fs.readFileSync(path.join(root, "src/taskpane.js"), "utf8");
  const imported = [...src.matchAll(/import\s+(\w+)\s+from\s+"\.\.\/csl\/([^"]+)"/g)];
  const style = imported.find(([, , f]) => f.endsWith(".csl"));
  const locales = {};
  for (const [, , f] of imported) {
    const m = /^locales-(.+)\.xml$/.exec(f);
    if (m) locales[m[1]] = fs.readFileSync(path.join(root, "csl", f), "utf8");
  }
  return createRenderer({ style: fs.readFileSync(path.join(root, "csl", style[2]), "utf8"), locales });
}

/** A small seeded random generator, so the same arguments give the same document. */
function rng(seed) {
  let s = seed >>> 0 || 1;
  return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32);
}

const FAMILY = ["Müller", "Schmidt", "Fischer", "Weber", "Smith", "Johnson", "Lee", "García", "Martin", "Brown", "Nguyen", "Kowalski", "Rossi", "Dubois", "Tanaka", "Andersen", "Novak", "Silva", "Okafor", "Haddad"];
const GIVEN = ["Anna", "Jan", "Lena", "Paul", "Maria", "Tom", "Eva", "Karl", "Sofia", "Ben", "Mia", "Noah"];
const JOURNAL = ["J Ecol", "Glob Change Biol", "Quat Sci Rev", "Clim Dyn", "Holocene", "J Glaciol", "Earth Surf Process Landf", "Geomorphology", "Ecol Lett", "Methods Ecol Evol"];
const WORDS = ("the of and in to a is that for on with as by at from this be are was it an or which were not have has its can these "
  + "glacier summer temperature tree ring record site sample snow depth survey model data analysis result change trend "
  + "valley slope basin core sediment climate season growth response effect estimate method field series region station").split(" ");
const TITLE = ("tree ring evidence of summer temperature in the alpine valley glacier retreat snow cover change since a regional "
  + "survey of long records from lake sediment cores and their response to climate variability over the last millennium").split(" ");

function library(n, rand) {
  const items = [];
  for (let i = 0; i < n; i++) {
    const author = Array.from({ length: 1 + Math.floor(rand() * 6) }, () => ({ family: FAMILY[Math.floor(rand() * FAMILY.length)], given: GIVEN[Math.floor(rand() * GIVEN.length)] }));
    const title = Array.from({ length: 8 + Math.floor(rand() * 8) }, () => TITLE[Math.floor(rand() * TITLE.length)]).join(" ");
    const y = 1995 + Math.floor(rand() * 31);
    const start = 1 + Math.floor(rand() * 2000);
    items.push({
      type: "article-journal", author, title: title[0].toUpperCase() + title.slice(1) + ` (${i})`,
      "container-title": JOURNAL[Math.floor(rand() * JOURNAL.length)], issued: { "date-parts": [[y]] },
      volume: String(1 + Math.floor(rand() * 400)), issue: String(1 + Math.floor(rand() * 12)), page: `${start}-${start + Math.floor(rand() * 15)}`,
      DOI: `10.9999/gen.${y}.${i}`,
    });
  }
  // The add-in's own merge gives the keys that an import would.
  return mergeItems({}, items).library;
}

function runXml(run) {
  const pr = [run.bold && "<w:b/>", run.italic && "<w:i/>", run.superscript && '<w:vertAlign w:val="superscript"/>', run.subscript && '<w:vertAlign w:val="subscript"/>'].filter(Boolean).join("");
  const rpr = pr ? `<w:rPr>${pr}</w:rPr>` : "";
  return run.text.split("\t").map((t, i) => (i ? `<w:r>${rpr}<w:tab/></w:r>` : "") + (t ? `<w:r>${rpr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : "")).join("");
}

/** Text of the heading above bibliography `n` (from 1). */
export const bibHeading = (n) => `References ${n}`;

/**
 * Builds the document. Returns {parts, store, stats}. A citation cites one
 * source (80 %), two (15 %) or three (5 %), now and then with a page; sources
 * are reused with a skew, and each section draws from a window of its own.
 */
export function generate({ pages = 20, citations = 100, bibs = 1, refs = 50, seed = 1, stale = false, paragraphsPerPage = 5, wordsPerParagraph = 95 } = {}) {
  const rand = rng(seed);
  const lib = library(refs, rand);
  const keys = Object.keys(lib);
  const paragraphs = pages * paragraphsPerPage;
  const at = new Map();
  for (let i = 0; i < citations; i++) {
    const p = Math.floor((i + 0.5) * paragraphs / citations);
    at.set(p, (at.get(p) ?? 0) + 1);
  }
  const perSection = Math.ceil(paragraphs / bibs);
  const stored = {};
  const controls = [];
  const body = [];
  let sdtId = 1000;
  for (let p = 0; p < paragraphs; p++) {
    const sec = Math.floor(p / perSection);
    if (p % perSection === 0) {
      if (p) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
      body.push(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${runXml({ text: `Section ${sec + 1}` })}</w:p>`);
    }
    const n = at.get(p) ?? 0;
    const words = Array.from({ length: wordsPerParagraph }, () => WORDS[Math.floor(rand() * WORDS.length)]);
    const chunks = [];
    const step = Math.floor(words.length / (n + 1));
    for (let k = 0; k <= n; k++) {
      chunks.push(runXml({ text: (k ? " " : "") + words.slice(k * step, k === n ? undefined : (k + 1) * step).join(" ") + (k < n ? " " : ".") }));
      if (k === n) break;
      const count = rand() < 0.8 ? 1 : rand() < 0.75 ? 2 : 3;
      const items = [];
      while (items.length < count) {
        const key = keys[(Math.floor(rand() ** 2 * Math.min(keys.length, 400)) + sec * 97) % keys.length];
        if (!items.some((x) => x.id === key)) items.push(rand() < 0.1 && count === 1 ? { id: key, locator: String(1 + Math.floor(rand() * 300)), label: "page" } : { id: key });
      }
      const id = newId() + controls.length.toString(36);
      stored[id] = items;
      controls.push({ kind: "cite", tag: CITE_PREFIX + id, title: citeTitle(items), current: CITE_NEW, original: CITE_NEW, sdt: sdtId++ });
      chunks.push({ cite: controls.length - 1 });
    }
    body.push(["<w:p>", ...chunks, "</w:p>"]);
    const last = p === paragraphs - 1 || (p + 1) % perSection === 0;
    if (last && sec < bibs) {
      body.push(`<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${runXml({ text: bibHeading(sec + 1) })}</w:p>`);
      controls.push({ kind: "bib", current: "", sdt: sdtId++ });
      body.push({ bib: controls.length - 1 });
    }
  }
  const locale = "en-US";
  const store = { ...emptyStore(locale), sources: lib, citations: stored };
  const renderer = nodeRenderer();
  const text = documentText(locale);
  const t0 = performance.now();
  const planned = plan(controls, store, renderer, { locale, bibEmpty: text.bibEmpty });
  const renderMs = performance.now() - t0;
  const citeRuns = new Map(planned.citations.map((c) => [c.index, c.runs]));
  const bibEntries = new Map(planned.bibs.map((b) => [b.index, b.entries]));

  const citeXml = (i) => {
    const c = controls[i];
    const runs = stale ? [{ text: CITE_NEW }] : citeRuns.get(i) ?? [{ text: CITE_NEW }];
    return `<w:sdt><w:sdtPr><w:alias w:val="${esc(c.title)}"/><w:tag w:val="${esc(c.tag)}"/><w:id w:val="${c.sdt}"/></w:sdtPr><w:sdtContent>${runs.map(runXml).join("")}</w:sdtContent></w:sdt>`;
  };
  const bibXml = (i) => {
    const entries = stale ? [[{ text: "…" }]] : bibEntries.get(i) ?? [[{ text: "…" }]];
    const ps = entries.map((runs) => `<w:p><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr>${runs.map(runXml).join("")}</w:p>`).join("");
    return `<w:sdt><w:sdtPr><w:alias w:val="${esc(text.bibTitle)}"/><w:tag w:val="${esc(BIB_TAG)}"/><w:id w:val="${controls[i].sdt}"/></w:sdtPr><w:sdtContent>${ps}</w:sdtContent></w:sdt>`;
  };
  const xml = body.map((b) => (typeof b === "string" ? b : Array.isArray(b) ? b.map((x) => (typeof x === "string" ? x : citeXml(x.cite))).join("") : bibXml(b.bib))).join("");

  const parts = { ...EMPTY };
  parts["word/document.xml"] = `${XML}<w:document ${W}><w:body>${xml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  parts["word/styles.xml"] = `${XML}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
    + `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>`
    + `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style></w:styles>`;
  Object.assign(parts, storeParts(store, seed));
  parts["word/_rels/document.xml.rels"] = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');

  const cited = new Set(Object.values(stored).flat().map((x) => x.id));
  return {
    parts,
    store,
    stats: {
      pages, paragraphs, citations: controls.filter((c) => c.kind === "cite").length, bibs: controls.filter((c) => c.kind === "bib").length,
      refs: keys.length, cited: cited.size,
      storeBytes: Buffer.byteLength(parts["customXml/item1.xml"]), documentBytes: Buffer.byteLength(parts["word/document.xml"]),
      nodeRenderMs: Math.round(renderMs), stale,
      // Sources that no citation uses yet: citing one at the start shifts every number.
      fresh: keys.filter((k) => !cited.has(k)).slice(0, 8),
    },
  };
}

/** The custom XML part that holds `store`, with its properties. Add the relationship to it yourself. */
export function storeParts(store, seed = 1) {
  return {
    "customXml/item1.xml": storeToXml(store),
    "customXml/itemProps1.xml": `${XML}<ds:datastoreItem ds:itemID="{6C3F2B1A-0000-4000-8000-${String(seed).padStart(12, "0")}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`,
    "customXml/_rels/item1.xml.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const opt = (n, d) => { const i = argv.indexOf("--" + n); return i >= 0 ? +argv[i + 1] : d; };
  const flags = new Set(["--stale", "--pane"]);
  const out = argv.find((a, i) => !a.startsWith("--") && !(i > 0 && argv[i - 1].startsWith("--") && !flags.has(argv[i - 1])));
  if (!out) { console.error("usage: node e2e/generate.mjs --pages N --citations C --bibs B --refs R [--stale] [--seed S] [--pane] out.docx"); process.exit(2); }
  const g = generate({ pages: opt("pages", 20), citations: opt("citations", 100), bibs: opt("bibs", 1), refs: opt("refs", 50), seed: opt("seed", 1), stale: argv.includes("--stale") });
  writeDocx(argv.includes("--pane") ? addPane(g.parts, idFor("measure.html")) : g.parts, path.resolve(out));
  console.log(JSON.stringify(g.stats));
}
