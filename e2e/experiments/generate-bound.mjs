// Documents for the binding experiment (e2e/experiments/README.md). Built on
// e2e/generate.mjs, written as OOXML directly.
//
//   node e2e/experiments/generate-bound.mjs --variant plain|rich|fields|unbound|features
//        [--pages 500 --citations 4000 --bibs 6 --refs 2000] [--xpath id|pos] [--pane] out.docx
//
// plain    every citation is a plain-text content control bound (w:dataBinding)
//          to one element of a custom XML part that holds every citation's text
// rich     the same with rich-text controls: the element holds the content as
//          escaped flat-OPC package XML, as Word 2013 and later store it
// fields   numeric citations as REF fields to bookmarked, numbered bibliography
//          paragraphs, each citation in a content control
// unbound  the generated document as it is, with the experiment's tags
// plain-unbound  plain-text controls with no binding (and the part, unused)
// features a small document: bound controls in the body, a footnote, a header,
//          a text box, a table cell and a comment, locked ones, a superscript
//          one, two bound to the same element, a rich one with italic text,
//          and a bookmark, a comment and a tracked insertion around citations
//
// Tags are "exp:cite:<id>" and "exp:bib" instead of the add-in's, so the real
// pane that runs next to the experiment page leaves the controls alone. The
// documents carry no store of the add-in. settings.xml asks Word to remove
// personal information on save.
import path from "node:path";
import { fileURLToPath } from "node:url";
import { generate } from "../generate.mjs";
import { XML, W, EMPTY, esc, addPane, writeDocx, idFor } from "../rig.mjs";
import { CITE_PREFIX, BIB_TAG } from "../../src/model.js";

import { NS, ITEM_ID, PREFIX, TAG, BIB, PAGE, xpathFor, storeXml, richPackage, runXml } from "./shared.js";
export { NS, ITEM_ID, TAG, BIB, PAGE };

// Rich-text controls bind through w15:dataBinding (Word 2013 and later); w:dataBinding is for plain text.
export const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const binding = (xpath, rich = false) => `<${rich ? "w15" : "w"}:dataBinding w:prefixMappings="${PREFIX}" w:xpath="${xpath}" w:storeItemID="${ITEM_ID}"/>`;
/** A bound control: plain-text (`<w:text/>`) or rich-text. */
export function boundSdt({ id, sdt, xpath, runs, rich = false, rPr = "", lock = "", alias = "Citation", bound = true }) {
  const content = runs.map((r) => runXml(r, rPr)).join("");
  return `<w:sdt><w:sdtPr>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:alias w:val="${esc(alias)}"/><w:tag w:val="${TAG}${esc(id)}"/><w:id w:val="${sdt}"/>${lock ? `<w:lock w:val="${lock}"/>` : ""}${bound ? binding(xpath, rich) : ""}${rich ? "" : "<w:text/>"}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
}

function storePartsFor(xml) {
  return {
    "customXml/item1.xml": xml,
    "customXml/itemProps1.xml": `${XML}<ds:datastoreItem ds:itemID="${ITEM_ID}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${NS}"/></ds:schemaRefs></ds:datastoreItem>`,
    "customXml/_rels/item1.xml.rels": `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
  };
}

const SETTINGS = `${XML}<w:settings ${W}><w:removePersonalInformation/><w:removeDateAndTime/></w:settings>`;
function addSettings(parts) {
  // w15 (Word 2013) for rich-text bindings, declared on the root and ignorable for older readers.
  for (const k of ["word/document.xml", "word/footnotes.xml", "word/header1.xml", "word/comments.xml"]) {
    if (parts[k]) parts[k] = parts[k].replace(/<w:(document|footnotes|hdr|comments) /, `<w:$1 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ${W15} mc:Ignorable="w15" `);
  }
  parts["word/settings.xml"] = SETTINGS;
  parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace("</Relationships>", '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>');
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>');
}

const CITE_SDT = /<w:sdt><w:sdtPr><w:alias w:val="([^"]*)"\/><w:tag w:val="siglum:cite:([^"]*)"\/><w:id w:val="(\d+)"\/><\/w:sdtPr><w:sdtContent>([\s\S]*?)<\/w:sdtContent><\/w:sdt>/g;
const plainOf = (runsXml) => [...runsXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&");

/**
 * The large documents. Returns {parts, stats, cites: [[id, text]]} where
 * `cites` lists the citations in document order with their current text.
 */
export function generateBound({ variant = "plain", xpath = "id", pages = 500, citations = 4000, bibs = 6, refs = 2000, seed = 1 } = {}) {
  const g = generate({ pages, citations, bibs, refs, seed });
  const parts = g.parts;
  let doc = parts["word/document.xml"];
  const cites = [];
  if (variant === "fields") return fieldsVariant(g);
  doc = doc.replace(CITE_SDT, (all, alias, id, sdt, content) => {
    const text = plainOf(content);
    const i = cites.length;
    cites.push([id, variant === "rich" ? richPackage(content) : text, text]);
    if (variant === "unbound") return `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${TAG}${id}"/><w:id w:val="${sdt}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
    // Plain-text binding needs one run; the citations here are plain "[n]" texts.
    return boundSdt({ id, sdt, xpath: xpathFor(id, i, xpath), runs: [{ text }], rich: variant === "rich", bound: variant !== "plain-unbound" });
  });
  doc = doc.replaceAll(`w:val="${BIB_TAG}"`, `w:val="${BIB}"`);
  parts["word/document.xml"] = doc;
  if (variant === "unbound") {
    delete parts["customXml/item1.xml"]; delete parts["customXml/itemProps1.xml"]; delete parts["customXml/_rels/item1.xml.rels"];
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(/<Relationship Id="rId2"[^>]*\/>/, "");
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(/<Override PartName="\/customXml\/itemProps1.xml"[^>]*\/>/, "");
  } else Object.assign(parts, storePartsFor(storeXml(cites.map(([id, v]) => [id, v]))));
  addSettings(parts);
  return { parts, cites: cites.map(([id, , text]) => [id, text]), stats: { ...g.stats, variant, xpath, storeBytes: Buffer.byteLength(parts["customXml/item1.xml"] ?? ""), documentBytes: Buffer.byteLength(doc) } };
}

/**
 * Candidate B: each bibliography is a numbered list (one list instance per
 * bibliography, restarting at 1), each entry bookmarked; each citation is
 * "[" REF \n \h fields joined by "," "]" in a content control.
 */
function fieldsVariant(g) {
  const parts = g.parts;
  let doc = parts["word/document.xml"];
  const store = g.store;
  // Number of each source per section, in order of first citation (citation-sequence).
  let section = 0;
  const order = [new Map()];
  const tokens = [];
  const re = /<w:tag w:val="siglum:cite:([^"]*)"\/>|<w:tag w:val="siglum:bibliography"\/>/g;
  for (const m of doc.matchAll(re)) {
    if (m[1] === undefined) { section++; order.push(new Map()); continue; }
    for (const it of store.citations[m[1]]) if (!order[section].has(it.id)) order[section].set(it.id, order[section].size);
    tokens.push(m[1]);
  }
  const bm = (sec, key) => `_Ref${String(sec + 1).padStart(2, "0")}${String(order[sec].get(key)).padStart(5, "0")}x`;
  let bmId = 1;
  section = 0;
  const cites = [];
  let fields = 0;
  // Citations first, bibliographies after, both in one pass over the body.
  doc = doc.replace(/(<w:sdt><w:sdtPr><w:alias w:val="([^"]*)"\/><w:tag w:val="siglum:cite:([^"]*)"\/><w:id w:val="(\d+)"\/><\/w:sdtPr><w:sdtContent>[\s\S]*?<\/w:sdtContent><\/w:sdt>)|<w:sdt><w:sdtPr><w:alias w:val="([^"]*)"\/><w:tag w:val="siglum:bibliography"\/><w:id w:val="(\d+)"\/><\/w:sdtPr><w:sdtContent>([\s\S]*?)<\/w:sdtContent><\/w:sdt>/g, (all, cite, alias, id, sdt, balias, bsdt, bcontent) => {
    if (cite) {
      const items = store.citations[id];
      const nums = items.map((it) => order[section].get(it.id) + 1);
      const refs = items.map((it, k) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ${bm(section, it.id)} \\n \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${nums[k]}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`);
      fields += refs.length;
      const loc = items.length === 1 && items[0].locator ? `<w:r><w:t xml:space="preserve">, p. ${esc(items[0].locator)}</w:t></w:r>` : "";
      cites.push([id, `[${nums.join(",")}${items.length === 1 && items[0].locator ? ", p. " + items[0].locator : ""}]`]);
      return `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${TAG}${id}"/><w:id w:val="${sdt}"/></w:sdtPr><w:sdtContent><w:r><w:t>[</w:t></w:r>${refs.join('<w:r><w:t>,</w:t></w:r>')}${loc}<w:r><w:t>]</w:t></w:r></w:sdtContent></w:sdt>`;
    }
    // A bibliography: the generated entries in citation order, numbered by the list.
    const entries = [...bcontent.matchAll(/<w:p>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
    const keys = [...order[section].keys()];
    if (entries.length !== keys.length) throw new Error(`bibliography ${section}: ${entries.length} entries, ${keys.length} sources`);
    const ps = entries.map((p, k) => {
      const key = keys[k];
      const title = store.sources[key].title;
      if (!plainOf(p).includes(title.slice(0, 30))) throw new Error(`entry ${k} of bibliography ${section} is not ${key}`);
      // Drop "n.\t" and the hanging indent; the list numbers the entry.
      const runs = p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "").replace(/^<w:p>/, "").replace(/<\/w:p>$/, "")
        .replace(/^<w:r><w:t xml:space="preserve">\d+\.<\/w:t><\/w:r><w:r><w:tab\/><\/w:r>/, "");
      const n = bmId++;
      return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${section + 1}"/></w:numPr></w:pPr><w:bookmarkStart w:id="${n}" w:name="${bm(section, key)}"/>${runs}<w:bookmarkEnd w:id="${n}"/></w:p>`;
    });
    section++;
    return `<w:sdt><w:sdtPr><w:alias w:val="${balias}"/><w:tag w:val="${BIB}"/><w:id w:val="${bsdt}"/></w:sdtPr><w:sdtContent>${ps.join("")}</w:sdtContent></w:sdt>`;
  });
  parts["word/document.xml"] = doc;
  delete parts["customXml/item1.xml"]; delete parts["customXml/itemProps1.xml"]; delete parts["customXml/_rels/item1.xml.rels"];
  parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(/<Relationship Id="rId2"[^>]*\/>/, "")
    .replace("</Relationships>", '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>');
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(/<Override PartName="\/customXml\/itemProps1.xml"[^>]*\/>/, "")
    .replace("</Types>", '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  const nums = order.slice(0, section).map((_, k) => `<w:num w:numId="${k + 1}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join("");
  parts["word/numbering.xml"] = `${XML}<w:numbering ${W}><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl></w:abstractNum>${nums}</w:numbering>`;
  addSettings(parts);
  return { parts, cites, stats: { ...g.stats, variant: "fields", fields, bookmarks: bmId - 1, documentBytes: Buffer.byteLength(doc) } };
}

/**
 * The small document with a bound citation in every place a citation can be.
 * Element ids: e1…e12; e5 is bound twice; r1 is the rich one.
 */
export function generateFeatures() {
  let sdt = 500;
  const p = (...xs) => `<w:p>${xs.join("")}</w:p>`;
  const t = (s) => `<w:r><w:t xml:space="preserve">${esc(s)}</w:t></w:r>`;
  const b = (id, opts = {}) => boundSdt({ id: opts.tag ?? id, sdt: sdt++, xpath: xpathFor(id), runs: opts.runs ?? [{ text: `[${id}]` }], ...opts });
  const etal = [{ text: "Author A " }, { text: "et al.", italic: true }, { text: " 2020" }];
  const store = [["e1", "[e1]"], ["e2", "[e2]"], ["e3", "[e3]"], ["e4", "[e4]"], ["e5", "[e5]"], ["e6", "[e6]"], ["e7", "[e7]"], ["e8", "[e8]"], ["e9", "[e9]"], ["e10", "[e10]"], ["e11", "[e11]"], ["e12", "[e12]"],
    ["r1", richPackage(etal.map((r) => runXml(r)).join(""))]];
  const body = [
    p(t("Plain: the survey counted 40 pairs "), b("e1"), t(".")),
    p(t("Locked content: "), b("e2", { lock: "sdtContentLocked" }), t(" and locked control: "), b("e3", { lock: "sdtLocked" }), t(" and both: "), b("e11", { lock: "contentLocked" }), t(".")),
    p(t("Superscript: the nest count rose"), b("e4", { rPr: '<w:vertAlign w:val="superscript"/>' }), t(".")),
    p(t("Same element twice: "), b("e5", { tag: "e5a" }), t(" and again "), b("e5", { tag: "e5b" }), t(".")),
    p(t("Rich with italic: "), b("r1", { rich: true, runs: etal }), t(".")),
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(t("Site 1"))}</w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(t("Table cell "), b("e6"))}</w:tc></w:tr></w:tbl>`,
    p(t("Footnote here"), `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>`, t(".")),
    p(t("Text box: "), `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="tb1" type="#_x0000_t202" style="width:200pt;height:40pt"><v:textbox><w:txbxContent>${p(t("In a box "), b("e9"))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`),
    p(`<w:commentRangeStart w:id="0"/>`, t("Commented text "), b("e12"), `<w:commentRangeEnd w:id="0"/>`, `<w:r><w:commentReference w:id="0"/></w:r>`),
    p(`<w:bookmarkStart w:id="90" w:name="keepMe"/>`, t("Bookmarked "), b("e1", { tag: "e1c" }), `<w:bookmarkEnd w:id="90"/>`, `<w:ins w:id="91" w:author="Author B" w:date="2026-01-01T00:00:00Z">${t(" tracked insertion")}</w:ins>`, t(".")),
  ];
  for (let k = 0; k < 40; k++) body.push(p(t("Filler paragraph about the wetland bird survey, counts of waders and gulls along the shore. ".repeat(4))));
  const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
  const parts = { ...EMPTY };
  parts["word/document.xml"] = `${XML}<w:document ${W}><w:body>${body.join("")}${sect}</w:body></w:document>`;
  parts["word/footnotes.xml"] = `${XML}<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${p(`<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>`, t(" In a footnote "), b("e7"))}</w:footnote></w:footnotes>`;
  parts["word/header1.xml"] = `${XML}<w:hdr ${W}>${p(t("Header "), b("e8"))}</w:hdr>`;
  parts["word/comments.xml"] = `${XML}<w:comments ${W}><w:comment w:id="0" w:author="Author B" w:initials="AB" w:date="2026-01-01T00:00:00Z">${p(t("Comment with a citation "), b("e10"))}</w:comment></w:comments>`;
  parts["word/_rels/document.xml.rels"] = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/><Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdCm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  Object.assign(parts, storePartsFor(storeXml(store)));
  addSettings(parts);
  return { parts, cites: store, stats: { variant: "features" } };
}

/** Four controls side by side: rich and plain, unbound and bound. Does Office.js see each? */
export function generateMini() {
  const t = (x) => `<w:r><w:t xml:space="preserve">${esc(x)}</w:t></w:r>`;
  const plain = (tag, extra) => `<w:sdt><w:sdtPr><w:tag w:val="${TAG}${tag}"/><w:id w:val="${700 + tag.length}"/>${extra}</w:sdtPr><w:sdtContent>${t("[" + tag + "]")}</w:sdtContent></w:sdt>`;
  const body = `<w:p>${t("Rich unbound ")}${plain("ru", "")}${t(" plain unbound ")}${plain("pu", "<w:text/>")}${t(" plain bound ")}${boundSdt({ id: "e1", sdt: 801, xpath: xpathFor("e1"), runs: [{ text: "[e1]" }] })}${process.env.MINI_RICH === "0" ? "" : t(" rich bound ") + boundSdt({ id: "r1", sdt: 802, xpath: xpathFor("r1"), runs: [{ text: "Author A et al." }], rich: true })}</w:p>`;
  const parts = { ...EMPTY };
  parts["word/document.xml"] = `${XML}<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
  parts["word/_rels/document.xml.rels"] = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  Object.assign(parts, storePartsFor(storeXml([["e1", "[e1 from part]"], ["r1", richPackage(runXml({ text: "Author B " }) + runXml({ text: "et al.", italic: true }))]])));
  addSettings(parts);
  return { parts, cites: [["e1", "[e1 from part]"]], stats: { variant: "mini" } };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const opt = (n, d) => { const i = argv.indexOf("--" + n); return i >= 0 ? argv[i + 1] : d; };
  const out = argv.at(-1);
  const variant = opt("variant", "plain");
  const g = variant === "features" ? generateFeatures() : generateBound({ variant, xpath: opt("xpath", "id"), pages: +opt("pages", 500), citations: +opt("citations", 4000), bibs: +opt("bibs", 6), refs: +opt("refs", 2000) });
  writeDocx(argv.includes("--pane") ? addPane(g.parts, idFor(PAGE)) : g.parts, path.resolve(out));
  console.log(JSON.stringify(g.stats));
}
