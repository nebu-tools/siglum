// Probes what a code review can only guess at, in real Word, with e2e/probe.js
// in the pane. Prints findings, not pass/fail: each line says what Word and
// the add-in did.
//
//   node e2e/probe.mjs
//
//   P1 places   citations in body, table, footnote, endnote, header and text
//               box; copy, cut and move between sections; a move with track
//               changes on; then saved.
//   S  saved    what the saved file holds: which citations were numbered, the
//               custom XML parts, stored citations against controls.
//   P2 reopen   P1's result opened again: does the first refresh find it current.
//   P3 paste    citations from P1 pasted into another document with another library.
//   CO two Word instances on one file (no pane): what the second one gets.
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { dist, build, requireWindows, inWord, writeDocx, readDocx, addPane, idFor, stop, EMPTY, PS } from "./rig.mjs";
import * as names from "./names.mjs";

requireWindows();
build();

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NSS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// Works with long keys, so a citation of five of them overflows the title.
const work = (id, family, year, title) => ({ id, type: "article-journal", author: [{ family, given: "A" }], title, "container-title": "Test J", issued: { "date-parts": [[year]] }, volume: "1", page: "1-9" });
const LIB = {
  alphaeins2020: work("alphaeins2020", "Alpha", 2020, "First work"),
  betazwei2019: work("betazwei2019", "Beta", 2019, "Second work"),
  gammadrei2018: work("gammadrei2018", "Gamma", 2018, "Third work"),
  deltavier2017: work("deltavier2017", "Delta", 2017, "Fourth work"),
  epsilonfuenf2016: work("epsilonfuenf2016", "Epsilon", 2016, "Fifth work"),
  zetasechs2015: work("zetasechs2015", "Zeta", 2015, "Sixth work in a footnote"),
  etasieben2014: work("etasieben2014", "Eta", 2014, "Seventh work in an endnote"),
  thetaacht2013: work("thetaacht2013", "Theta", 2013, "Eighth work in a header"),
  iotaneun2012: work("iotaneun2012", "Iota", 2012, "Ninth work in a text box"),
};

let sdt = 500;
const citations = {};
function cite(label, items) {
  const id = names.newId() + label;
  citations[id] = items;
  return `<w:sdt><w:sdtPr><w:alias w:val="${esc(names.citeTitle(items))}"/><w:tag w:val="${esc(names.citeTag(id))}"/><w:id w:val="${sdt++}"/></w:sdtPr><w:sdtContent><w:r><w:t>${esc(names.CITE_NEW)}</w:t></w:r></w:sdtContent></w:sdt>`;
}
const run = (t) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const p = (...xs) => `<w:p>${xs.join("")}</w:p>`;
const bib = () => `<w:sdt><w:sdtPr><w:alias w:val="Literaturverzeichnis"/><w:tag w:val="${esc(names.BIB_TAG)}"/><w:id w:val="${sdt++}"/></w:sdtPr><w:sdtContent>${p(run("…"))}</w:sdtContent></w:sdt>`;
const one = (k) => [{ id: k }];

function placesDoc() {
  const parts = { ...EMPTY };
  const body = [
    p(run("Abschnitt 1")),
    p(run("Body one "), cite("A", one("alphaeins2020")), run(".")),
    p(run("Body two "), cite("B", [{ id: "betazwei2019", locator: "12", label: "page" }]), run(".")),
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(run("Cell "), cite("C", one("gammadrei2018")))}</w:tc></w:tr></w:tbl>`,
    p(run("Footnote here"), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', run(" and endnote here"), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>', run(".")),
    p(run("Text box: "), `<w:r><w:pict><v:shape id="tb1" type="#_x0000_t202" style="width:220pt;height:40pt"><v:textbox><w:txbxContent>${p(run("In box "), cite("TB", one("iotaneun2012")))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`),
    p(run("Referenzliste 1")), bib(),
    p(run("Abschnitt 2")),
    p(run("Section two "), cite("S2", one("deltavier2017")), run(".")),
    p(run("Ziel")),
    p(run("Referenzliste 2")), bib(),
    p(run("Ende")),
  ];
  const sep = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  parts["word/footnotes.xml"] = `${X}<w:footnotes ${NSS}>${sep}<w:footnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>', run(" In a footnote "), cite("FN", one("zetasechs2015")))}</w:footnote></w:footnotes>`;
  parts["word/endnotes.xml"] = `${X}<w:endnotes ${NSS}>${sep.replaceAll("footnote", "endnote")}<w:endnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>', run(" In an endnote "), cite("EN", one("etasieben2014")))}</w:endnote></w:endnotes>`;
  parts["word/header1.xml"] = `${X}<w:hdr ${NSS}>${p(run("Header "), cite("HD", one("thetaacht2013")))}</w:hdr>`;
  parts["word/document.xml"] = `${X}<w:document ${NSS}><w:body>${body.join("")}<w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr></w:body></w:document>`;
  parts["word/_rels/document.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>'
    + '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>';
  parts["customXml/item1.xml"] = names.storeToXml(names.makeStore(LIB, citations));
  parts["customXml/itemProps1.xml"] = `${X}<ds:datastoreItem ds:itemID="{6C3F2B1A-0000-4000-8000-0000000000AA}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`;
  parts["customXml/_rels/item1.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", [
    ["/word/footnotes.xml", "footnotes"], ["/word/endnotes.xml", "endnotes"], ["/word/header1.xml", "header"],
  ].map(([n, t]) => `<Override PartName="${n}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${t}+xml"/>`).join("")
    + '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  return parts;
}

/** Another document whose library has only some of P1's works. */
function otherDoc() {
  const parts = { ...EMPTY };
  const lib = { alphaeins2020: LIB.alphaeins2020, betazwei2019: LIB.betazwei2019, gammadrei2018: LIB.gammadrei2018, deltavier2017: LIB.deltavier2017, epsilonfuenf2016: LIB.epsilonfuenf2016 };
  parts["word/document.xml"] = `${X}<w:document ${NSS}><w:body>${[p(run("Other document")), p(run("Paste one")), p(run("Paste two")), p(run("Paste three")), p(run("Paste four")), p(run("Referenzliste")), bib()].join("")}<w:sectPr/></w:body></w:document>`;
  parts["word/_rels/document.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>`;
  parts["customXml/item1.xml"] = names.storeToXml(names.makeStore(lib, {}));
  parts["customXml/itemProps1.xml"] = `${X}<ds:datastoreItem ds:itemID="{6C3F2B1A-0000-4000-8000-0000000000BB}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`;
  parts["customXml/_rels/item1.xml.rels"] = `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`;
  parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace("</Types>", '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>');
  return parts;
}

const show = (m) => {
  // P1 hands its citations' OOXML to P3 through the served directory.
  if (m.clips) { fs.writeFileSync(path.join(dist, "probe-clips.json"), JSON.stringify(m.clips)); return; }
  if (m.probe) console.log(`  ${m.probe}: ${JSON.stringify(m.data)}`);
  else if (!m.done) console.log("    " + JSON.stringify(m));
};
const config = (o) => fs.writeFileSync(path.join(dist, "probe-config.json"), JSON.stringify(o));
const id = idFor("probe.html");
const file = (n) => path.join(dist, n);

/** Citation controls in a saved part: [tag, text]. */
function sdts(xml) {
  const out = [];
  for (const m of (xml ?? "").matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)) {
    const tag = /<w:tag w:val="([^"]*)"/.exec(m[1])?.[1];
    if (!names.isCiteTag(tag)) continue;
    const content = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>/.exec(m[1])?.[1] ?? "";
    out.push([tag.slice(names.CITE_PREFIX.length), [...content.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join("")]);
  }
  return out;
}

try {
  console.log("\n== P1 places, copy, move, tracked move");
  writeDocx(addPane(placesDoc(), id), file("probe-p1-in.docx"));
  config({ mode: "places" });
  let r = await inWord({ docx: file("probe-p1-in.docx"), out: file("probe-p1.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== S what the saved file holds");
  const saved = readDocx(file("probe-p1.docx"));
  const textbox = [...(saved["word/document.xml"] ?? "").matchAll(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/g)].map((m) => m[1]).join("");
  for (const [part, xml] of [["footnote", saved["word/footnotes.xml"]], ["endnote", saved["word/endnotes.xml"]], ["header", Object.entries(saved).filter(([k]) => /^word\/header\d*\.xml$/.test(k)).map(([, v]) => v).join("")], ["text box", textbox]]) {
    console.log(`  citation in ${part}: ${JSON.stringify(sdts(xml))}`);
  }
  const stores = Object.entries(saved).filter(([k, v]) => /^customXml\/item\d+\.xml$/.test(k) && typeof v === "string" && v.includes(names.NS));
  console.log(`  custom XML parts with ${names.NS}: ${stores.length}`);
  const st = names.storeFromXml(stores[0]?.[1]);
  const tagsInDoc = new Set(Object.values(saved).filter((v) => typeof v === "string").flatMap((v) => sdts(v).map(([t]) => t)));
  const stored = Object.keys(names.citationsOf(st));
  console.log(`  stored citations ${stored.length}, citation controls in the file ${tagsInDoc.size}, stored without a control ${stored.filter((t) => !tagsInDoc.has(t)).length}, controls without a stored entry ${[...tagsInDoc].filter((t) => !names.citationsOf(st)[t]).length}`);
  console.log(`  works in the library ${Object.keys(names.libraryOf(st)).length}`);

  console.log("\n== P2 reopen the saved file");
  config({ mode: "reopen" });
  const reopenIn = file("probe-p2-in.docx");
  fs.copyFileSync(file("probe-p1.docx"), reopenIn);
  r = await inWord({ docx: reopenIn, out: file("probe-p2.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== P3 paste into another document");
  config({ mode: "paste" });
  writeDocx(addPane(otherDoc(), id), file("probe-p3-in.docx"));
  r = await inWord({ docx: file("probe-p3-in.docx"), out: file("probe-p3.docx"), page: "probe.html", onMessage: show });
  console.log(`  (Word exit ${r.code}, saved ${r.saved})`);

  console.log("\n== CO two Word instances on one file, no pane");
  const winTemp = execFileSync(PS, ["-NoProfile", "-Command", "$env:TEMP"], { encoding: "utf8" }).trim();
  const winDir = `${winTemp}\\siglum-co-${process.pid}`;
  const coDir = execFileSync("wslpath", ["-u", winDir], { encoding: "utf8" }).trim();
  fs.mkdirSync(coDir, { recursive: true });
  fs.copyFileSync(file("probe-p1.docx"), path.join(coDir, "shared.docx"));
  fs.copyFileSync(path.join(path.dirname(new URL(import.meta.url).pathname), "coauthor.ps1"), path.join(coDir, "coauthor.ps1"));
  try {
    const out = execFileSync(PS, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", `${winDir}\\coauthor.ps1`, "-File", `${winDir}\\shared.docx`], { encoding: "utf8" });
    console.log(out.replace(/\r/g, "").trim().split("\n").map((l) => "  " + l).join("\n"));
  } finally {
    fs.rmSync(coDir, { recursive: true, force: true });
  }
} finally {
  stop();
}
