// Generates a large test document: N pages of text in B sections, C citations
// spread through them, a reference list at the end of each section and a
// library of R works, all in the add-in's own stored form (names from
// e2e/names.mjs). Citation numbers and lists are rendered with the add-in's
// own plan() and citeproc, so the document is current when Word opens it;
// with --stale every citation reads as a new one and every list is empty.
//
//   node e2e/generate.mjs --pages 20 --citations 100 --bibs 1 --refs 50 [--stale] [--seed 1] [--pane] out.docx
//
// --pane adds a task pane for the measuring page that opens with the document.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRenderer } from "../src/render.js";
import { mergeItems } from "../src/library.js";
import * as names from "./names.mjs";
import { root, EMPTY, addPane, writeDocx, idFor } from "./rig.mjs";

/** The add-in's renderer, with the style, locales and language taskpane.js uses. */
export function nodeRenderer() {
  const src = fs.readFileSync(path.join(root, "src/taskpane.js"), "utf8");
  const styleFile = /from\s+"\.\.\/csl\/([^"]+\.csl)"/.exec(src)?.[1] ?? fs.readdirSync(path.join(root, "csl")).find((f) => f.endsWith(".csl"));
  const lang = /lang:\s*"([a-z]{2}-[A-Z]{2})"/.exec(src)?.[1] ?? "de-DE";
  const locales = {};
  for (const f of fs.readdirSync(path.join(root, "csl"))) {
    const m = /^locales-(.+)\.xml$/.exec(f);
    if (m) locales[m[1]] = fs.readFileSync(path.join(root, "csl", f), "utf8");
  }
  return createRenderer({ style: fs.readFileSync(path.join(root, "csl", styleFile), "utf8"), locales, lang });
}

/** A small seeded PRNG, so the same arguments give the same document. */
function rng(seed) {
  let s = seed >>> 0 || 1;
  return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 2 ** 32);
}

const FAMILY = ["Müller", "Schmidt", "Schneider", "Fischer", "Weber", "Meyer", "Wagner", "Becker", "Schulz", "Hoffmann", "Smith", "Johnson", "Lee", "Garcia", "Martin", "Brown", "Nguyen", "Kowalski", "Rossi", "Dubois", "Tanaka", "Andersen", "Novak", "Silva"];
const GIVEN = ["Anna", "Jan", "Lena", "Paul", "Maria", "Tom", "Eva", "Karl", "Sofia", "Ben", "Mia", "Noah"];
const JOURNAL = ["N Engl J Med", "Lancet", "JAMA", "BMJ", "Ann Oncol", "J Clin Oncol", "Dtsch Arztebl Int", "Lancet Oncol", "Eur J Cancer", "BMC Med Res Methodol"];
const WORDS = ("die der und in zu den das nicht von sie ist des sich mit dem dass er es ein ich auf so eine auch als an nach wie im für "
  + "Studie Patienten Behandlung Wirksamkeit Sicherheit Endpunkt Analyse Ergebnisse Daten Therapie Vergleich randomisiert kontrolliert "
  + "Überlebenszeit Nebenwirkungen Population Subgruppe signifikant Effekt Hazard Ratio Konfidenzintervall Nutzen Bewertung Dossier "
  + "Zusatznutzen Vergleichstherapie Morbidität Lebensqualität Mortalität Studienarm Placebo Dosis Woche Monate Zulassung").split(" ");
const TITLE = ("overall survival progression free randomized controlled trial phase patients with advanced cancer treatment efficacy "
  + "safety of versus in a multicenter study outcomes quality life meta analysis systematic review adjuvant therapy").split(" ");

function library(n, rand) {
  const items = [];
  for (let i = 0; i < n; i++) {
    const authors = Array.from({ length: 1 + Math.floor(rand() * 6) }, () => ({ family: FAMILY[Math.floor(rand() * FAMILY.length)], given: GIVEN[Math.floor(rand() * GIVEN.length)] }));
    const title = Array.from({ length: 8 + Math.floor(rand() * 8) }, () => TITLE[Math.floor(rand() * TITLE.length)]).join(" ");
    const y = 1995 + Math.floor(rand() * 31);
    const start = 1 + Math.floor(rand() * 2000);
    items.push({
      type: "article-journal", author: authors, title: title[0].toUpperCase() + title.slice(1) + ` (${i})`,
      "container-title": JOURNAL[Math.floor(rand() * JOURNAL.length)], issued: { "date-parts": [[y]] },
      volume: String(1 + Math.floor(rand() * 400)), issue: String(1 + Math.floor(rand() * 12)), page: `${start}-${start + Math.floor(rand() * 15)}`,
      DOI: `10.9999/gen.${y}.${i}`, PMID: String(30000000 + i),
    });
  }
  // The add-in's own merge gives the keys a real import would.
  return mergeItems({}, items).library;
}

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function runXml(run) {
  const pr = [run.bold && "<w:b/>", run.italic && "<w:i/>", run.superscript && '<w:vertAlign w:val="superscript"/>', run.subscript && '<w:vertAlign w:val="subscript"/>'].filter(Boolean).join("");
  const rpr = pr ? `<w:rPr>${pr}</w:rPr>` : "";
  return run.text.split("\t").map((t, i) => (i ? `<w:r>${rpr}<w:tab/></w:r>` : "") + (t ? `<w:r>${rpr}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : "")).join("");
}

/**
 * Builds the document. Returns {parts, stats}: the docx parts and what went in.
 * Citations cite one work (80 %), two (15 %) or three (5 %), with a page
 * locator now and then; works are reused with a skew, as in real dossiers.
 */
export function generate({ pages = 20, citations = 100, bibs = 1, refs = 50, seed = 1, stale = false, paragraphsPerPage = 5, wordsPerParagraph = 95 } = {}) {
  const rand = rng(seed);
  const lib = library(refs, rand);
  const keys = Object.keys(lib);
  const paragraphs = pages * paragraphsPerPage;
  // Where each citation goes: evenly spread over the body paragraphs.
  const at = new Map();
  for (let i = 0; i < citations; i++) {
    const p = Math.floor((i + 0.5) * paragraphs / citations);
    at.set(p, (at.get(p) ?? 0) + 1);
  }
  const perSection = Math.ceil(paragraphs / bibs);
  const zitate = {};
  const controls = [];
  const body = []; // strings, with {cite: index} placeholders for citation controls
  let sdtId = 1000;
  let cites = 0;
  for (let p = 0; p < paragraphs; p++) {
    const sec = Math.floor(p / perSection);
    if (p % perSection === 0) {
      if (p) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
      body.push(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${runXml({ text: `Abschnitt ${sec + 1}` })}</w:p>`);
    }
    const n = at.get(p) ?? 0;
    const words = Array.from({ length: wordsPerParagraph }, () => WORDS[Math.floor(rand() * WORDS.length)]);
    const chunks = [];
    const step = Math.floor(words.length / (n + 1));
    for (let k = 0; k <= n; k++) {
      chunks.push(runXml({ text: (k ? " " : "") + words.slice(k * step, k === n ? undefined : (k + 1) * step).join(" ") + (k < n ? " " : ".") }));
      if (k === n) break;
      // A citation: works from a section-specific window of the library, skewed.
      const count = rand() < 0.8 ? 1 : rand() < 0.75 ? 2 : 3;
      const items = [];
      while (items.length < count) {
        const k2 = keys[(Math.floor(rand() ** 2 * Math.min(keys.length, 400)) + sec * 97) % keys.length];
        if (!items.some((x) => x.id === k2)) items.push(rand() < 0.1 && count === 1 ? { id: k2, locator: String(1 + Math.floor(rand() * 300)), label: "page" } : { id: k2 });
      }
      const id = names.newId() + cites.toString(36);
      zitate[id] = items;
      controls.push({ kind: "cite", tag: names.citeTag(id), title: names.citeTitle(items), current: names.CITE_NEW, original: names.CITE_NEW, sdt: sdtId++ });
      chunks.push({ cite: controls.length - 1 });
      cites++;
    }
    body.push(["<w:p>", ...chunks, "</w:p>"]);
    const last = p === paragraphs - 1 || (p + 1) % perSection === 0;
    if (last && sec < bibs) {
      body.push(`<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${runXml({ text: `Referenzliste ${sec + 1}` })}</w:p>`);
      controls.push({ kind: "bib", current: "", sdt: sdtId++ });
      body.push({ bib: controls.length - 1 });
    }
  }
  const store = names.makeStore(lib, zitate);
  const renderer = nodeRenderer();
  const t0 = performance.now();
  const plan = names.plan(controls, store, renderer);
  const renderMs = performance.now() - t0;
  const citeRuns = new Map(plan.citations.map((c) => [c.index, c.runs]));
  const bibEntries = new Map(plan.bibs.map((b) => [b.index, b.entries]));

  const citeXml = (i) => {
    const c = controls[i];
    const runs = stale ? [{ text: names.CITE_NEW }] : citeRuns.get(i) ?? [{ text: names.CITE_NEW }];
    return `<w:sdt><w:sdtPr><w:alias w:val="${esc(c.title)}"/><w:tag w:val="${esc(c.tag)}"/><w:id w:val="${c.sdt}"/></w:sdtPr><w:sdtContent>${runs.map(runXml).join("")}</w:sdtContent></w:sdt>`;
  };
  const bibXml = (i) => {
    const entries = stale ? [[{ text: "…" }]] : bibEntries.get(i) ?? [[{ text: "…" }]];
    const ps = entries.map((runs) => `<w:p><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr>${runs.map(runXml).join("")}</w:p>`).join("");
    return `<w:sdt><w:sdtPr><w:alias w:val="Literaturverzeichnis"/><w:tag w:val="${esc(names.BIB_TAG)}"/><w:id w:val="${controls[i].sdt}"/></w:sdtPr><w:sdtContent>${ps}</w:sdtContent></w:sdt>`;
  };
  const xml = body.map((b) => (typeof b === "string" ? b : Array.isArray(b) ? b.map((x) => (typeof x === "string" ? x : citeXml(x.cite))).join("") : bibXml(b.bib))).join("");

  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const parts = { ...EMPTY };
  parts["word/document.xml"] = `${X}<w:document ${W}><w:body>${xml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  parts["word/styles.xml"] = `${X}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
    + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
    + `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>`
    + `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style></w:styles>`;
  parts["word/_rels/document.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`;
  parts["customXml/item1.xml"] = names.storeToXml(store);
  parts["customXml/itemProps1.xml"] = `${X}<ds:datastoreItem ds:itemID="{6C3F2B1A-0000-4000-8000-${String(seed).padStart(12, "0")}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`;
  parts["customXml/_rels/item1.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');

  const cited = new Set(Object.values(zitate).flat().map((x) => x.id));
  return {
    parts,
    store,
    stats: {
      pages, paragraphs, citations: cites, bibs: controls.filter((c) => c.kind === "bib").length, refs: keys.length, cited: cited.size,
      storeBytes: Buffer.byteLength(parts["customXml/item1.xml"]), documentBytes: Buffer.byteLength(parts["word/document.xml"]),
      nodeRenderMs: Math.round(renderMs), stale,
      // Works no citation uses yet: citing one at the start shifts every number.
      fresh: keys.filter((k) => !cited.has(k)).slice(0, 8),
    },
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const opt = (n, d) => { const i = argv.indexOf("--" + n); return i >= 0 ? +argv[i + 1] : d; };
  const out = argv.filter((a, i) => !a.startsWith("--") && !(i > 0 && argv[i - 1].startsWith("--") && !["--stale", "--pane"].includes(argv[i - 1])))[0];
  if (!out) { console.error("usage: node e2e/generate.mjs --pages N --citations C --bibs B --refs R [--stale] [--seed S] [--pane] out.docx"); process.exit(2); }
  const g = generate({ pages: opt("pages", 20), citations: opt("citations", 100), bibs: opt("bibs", 1), refs: opt("refs", 50), seed: opt("seed", 1), stale: argv.includes("--stale") });
  writeDocx(argv.includes("--pane") ? addPane(g.parts, idFor("measure.html")) : g.parts, path.resolve(out));
  console.log(JSON.stringify(g.stats));
}
