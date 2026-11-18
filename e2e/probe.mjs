// Probes in real Word what the README lists as limits, with e2e/probe.js in
// the pane. It prints findings, not pass or fail: each line says what Word
// and the add-in did.
//
//   node e2e/probe.mjs
//
//   P1 places   citations in the body, a table, a footnote, an endnote, a
//               header and a text box, and one with a title cut at 64 characters
//   S  saved    what the saved file holds in each of those places
//   P2 reopen   P1's result opened again
//   P3 paste    P1's citations pasted into another document
//   CO          two Word instances on one file, without the pane
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { root, dist, build, requireWindows, inWord, writeDocx, readDocx, addPane, idFor, stop, EMPTY, XML, esc, PS } from "./rig.mjs";
import { storeParts } from "./generate.mjs";
import { CITE_PREFIX, BIB_TAG, CITE_NEW, NS, emptyStore, storeFromXml, citeTitle, newId } from "../src/model.js";

requireWindows();
build();

const NSS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';

// Sources with long keys, so a citation of five of them overflows the title.
const source = (id, family, title) => ({ id, type: "article-journal", author: [{ family, given: "A" }], title, "container-title": "Test J", issued: { "date-parts": [[2020]] }, volume: "1", page: "1-9" });
const SOURCES = {
  alphaone2020: source("alphaone2020", "Alpha", "First work"),
  betatwo2019: source("betatwo2019", "Beta", "Second work"),
  gammathree2018: source("gammathree2018", "Gamma", "Third work"),
  deltafour2017: source("deltafour2017", "Delta", "Fourth work"),
  epsilonfive2016: source("epsilonfive2016", "Epsilon", "Fifth work"),
  zetasix2015: source("zetasix2015", "Zeta", "Sixth work, in a footnote"),
  etaseven2014: source("etaseven2014", "Eta", "Seventh work, in an endnote"),
  thetaeight2013: source("thetaeight2013", "Theta", "Eighth work, in a header"),
  iotanine2012: source("iotanine2012", "Iota", "Ninth work, in a text box"),
};

let sdt = 500;
const stored = {};
function cite(label, items) {
  const id = newId() + label;
  stored[id] = items;
  return `<w:sdt><w:sdtPr><w:alias w:val="${esc(citeTitle(items))}"/><w:tag w:val="${esc(CITE_PREFIX + id)}"/><w:id w:val="${sdt++}"/></w:sdtPr><w:sdtContent><w:r><w:t>${esc(CITE_NEW)}</w:t></w:r></w:sdtContent></w:sdt>`;
}
const run = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const p = (...xs) => `<w:p>${xs.join("")}</w:p>`;
const bib = () => `<w:sdt><w:sdtPr><w:alias w:val="Bibliography"/><w:tag w:val="${esc(BIB_TAG)}"/><w:id w:val="${sdt++}"/></w:sdtPr><w:sdtContent>${p(run("…"))}</w:sdtContent></w:sdt>`;
const one = (k) => [{ id: k }];
const withStore = (parts, store, rels, seed) => {
  Object.assign(parts, storeParts(store, seed));
  parts["word/_rels/document.xml.rels"] = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}<Relationship Id="rIdStore" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  return parts;
};

function placesDoc() {
  const parts = { ...EMPTY };
  const body = [
    p(run("Section 1")),
    p(run("Body one "), cite("A", one("alphaone2020")), run(".")),
    p(run("Body two "), cite("B", [{ id: "betatwo2019", locator: "12", label: "page" }]), run(".")),
    p(run("Five sources "), cite("LONG", ["alphaone2020", "betatwo2019", "gammathree2018", "deltafour2017", "epsilonfive2016"].map((id) => ({ id }))), run(".")),
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(run("Cell "), cite("C", one("gammathree2018")))}</w:tc></w:tr></w:tbl>`,
    p(run("Footnote here"), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', run(" and endnote here"), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>', run(".")),
    p(run("Text box: "), `<w:r><w:pict><v:shape id="tb1" type="#_x0000_t202" style="width:220pt;height:40pt"><v:textbox><w:txbxContent>${p(run("In a box "), cite("TB", one("iotanine2012")))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`),
    p(run("References 1")), bib(),
  ];
  const sep = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  parts["word/footnotes.xml"] = `${XML}<w:footnotes ${NSS}>${sep}<w:footnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>', run(" In a footnote "), cite("FN", one("zetasix2015")))}</w:footnote></w:footnotes>`;
  parts["word/endnotes.xml"] = `${XML}<w:endnotes ${NSS}>${sep.replaceAll("footnote", "endnote")}<w:endnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>', run(" In an endnote "), cite("EN", one("etaseven2014")))}</w:endnote></w:endnotes>`;
  parts["word/header1.xml"] = `${XML}<w:hdr ${NSS}>${p(run("Header "), cite("HD", one("thetaeight2013")))}</w:hdr>`;
  parts["word/document.xml"] = `${XML}<w:document ${NSS}><w:body>${body.join("")}<w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr></w:body></w:document>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", [
    ["/word/footnotes.xml", "footnotes"], ["/word/endnotes.xml", "endnotes"], ["/word/header1.xml", "header"],
  ].map(([n, t]) => `<Override PartName="${n}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${t}+xml"/>`).join("") + "</Types>");
  return withStore(parts, { ...emptyStore("en-US"), sources: SOURCES, citations: stored },
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
    + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>', 170);
}

/** Another document whose sources are only some of P1's. */
function otherDoc() {
  const parts = { ...EMPTY };
  const sources = Object.fromEntries(Object.entries(SOURCES).slice(0, 5));
  parts["word/document.xml"] = `${XML}<w:document ${NSS}><w:body>${[p(run("Other document")), p(run("Paste one")), p(run("Paste two")), p(run("Paste three")), p(run("References")), bib()].join("")}<w:sectPr/></w:body></w:document>`;
  return withStore(parts, { ...emptyStore("en-US"), sources }, "", 187);
}

const show = (m) => {
  if (m.clips) { fs.writeFileSync(path.join(dist, "probe-clips.json"), JSON.stringify(m.clips)); return; }
  if (m.probe) console.log(`  ${m.probe}: ${JSON.stringify(m.data)}`);
  else if (!m.done) console.log("    " + JSON.stringify(m));
};
const config = (o) => fs.writeFileSync(path.join(dist, "probe-config.json"), JSON.stringify(o));
const id = idFor("probe.html");
const file = (n) => path.join(dist, n);

/** Citation controls in a saved part: [label, text]. */
function sdts(xml) {
  const out = [];
  for (const m of (xml ?? "").matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)) {
    const tag = /<w:tag w:val="([^"]*)"/.exec(m[1])?.[1];
    if (!tag?.startsWith(CITE_PREFIX)) continue;
    const content = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>/.exec(m[1])?.[1] ?? "";
    out.push([tag.slice(-2), [...content.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("")]);
  }
  return out;
}

try {
  console.log("\n== P1 places");
  writeDocx(addPane(placesDoc(), id), file("probe-p1-in.docx"));
  config({ mode: "places" });
  let r = await inWord({ docx: file("probe-p1-in.docx"), out: file("probe-p1.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== S what the saved file holds");
  const saved = readDocx(file("probe-p1.docx"));
  const textbox = [...(saved["word/document.xml"] ?? "").matchAll(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g)].map((m) => m[1]).join("");
  for (const [part, xml] of [["footnote", saved["word/footnotes.xml"]], ["endnote", saved["word/endnotes.xml"]], ["header", Object.entries(saved).filter(([k]) => /^word\/header\d*\.xml$/.test(k)).map(([, v]) => v).join("")], ["text box", textbox]]) {
    console.log(`  citation in a ${part}: ${JSON.stringify(sdts(xml))}`);
  }
  const stores = Object.entries(saved).filter(([k, v]) => /^customXml\/item\d+\.xml$/.test(k) && typeof v === "string" && v.includes(NS));
  console.log(`  custom XML parts in ${NS}: ${stores.length}, sources ${Object.keys(storeFromXml(stores[0][1]).sources).length}`);

  console.log("\n== P2 reopen");
  config({ mode: "reopen" });
  writeDocx(addPane(readDocx(file("probe-p1.docx")), id), file("probe-p2-in.docx"));
  r = await inWord({ docx: file("probe-p2-in.docx"), out: file("probe-p2.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== P3 paste into another document");
  config({ mode: "paste" });
  writeDocx(addPane(otherDoc(), id), file("probe-p3-in.docx"));
  r = await inWord({ docx: file("probe-p3-in.docx"), out: file("probe-p3.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== CO two Word instances on one file, without the pane");
  const winTemp = execFileSync(PS, ["-NoProfile", "-Command", "$env:TEMP"], { encoding: "utf8" }).trim();
  const winDir = `${winTemp}\\siglum-co-${process.pid}`;
  const coDir = execFileSync("wslpath", ["-u", winDir], { encoding: "utf8" }).trim();
  fs.mkdirSync(coDir, { recursive: true });
  try {
    fs.copyFileSync(file("probe-p1.docx"), path.join(coDir, "shared.docx"));
    fs.copyFileSync(path.join(root, "e2e/coauthor.ps1"), path.join(coDir, "coauthor.ps1"));
    const out = execFileSync(PS, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", `${winDir}\\coauthor.ps1`, "-File", `${winDir}\\shared.docx`], { encoding: "utf8" });
    console.log(out.replace(/\r/g, "").trim().split("\n").map((l) => "  " + l).join("\n"));
  } finally {
    fs.rmSync(coDir, { recursive: true, force: true });
  }
} finally {
  stop();
}
