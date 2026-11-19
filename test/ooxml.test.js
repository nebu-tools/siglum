import { test } from "node:test";
import assert from "node:assert/strict";
import { controlTexts, paragraphStyle, OoxmlError, bookmarkNames, rewriteCitationTexts } from "../src/ooxml.js";
import { W, markup, pkg, bodyPackage } from "./ooxml-fixtures.js";

const { sdt, r, p, del, ins, pDeletedMark } = markup();
const texts = (x) => controlTexts(bodyPackage(x)).map((c) => [c.tag, c.current, c.original]);
const unreadable = (fn) => assert.throws(fn, (e) => e instanceof OoxmlError && e.key === "error.ooxml");

test("plain controls, in document order, nested ones too", () => {
  assert.deepEqual(texts(p(`${r("One ")}${sdt("a", r("[1]"))}${r(" and ")}${sdt("outer", r("x ") + sdt("b", r("[2]")))}`)), [["a", "[1]", "[1]"], ["outer", "x [2]", "x [2]"], ["b", "[2]", "[2]"]]);
});

test("a tracked deletion, an insertion and a replacement", () => {
  assert.deepEqual(texts(p(`${sdt("gone", del("[1]"))}${ins(sdt("new", r("[2]")))}${sdt("edit", del("[3]") + ins(r("[4]")))}`)), [["gone", "", "[1]"], ["new", "[2]", ""], ["edit", "[4]", "[3]"]]);
});

test("a control inside a deletion is only in the original text", () => {
  const deleted = `<w:del w:id="7" w:author="A">${sdt("in", `<w:r><w:delText>[5]</w:delText></w:r>`)}</w:del>`;
  assert.deepEqual(texts(p(r("x") + deleted)), [["in", "", "[5]"]]);
});

test("a control nested in another, with a change in the inner one", () => {
  assert.deepEqual(texts(p(sdt("outer", r("see ") + sdt("inner", del("[1]") + ins(r("[2]")))))), [["outer", "see [2]", "see [1]"], ["inner", "[2]", "[1]"]]);
});

test("tracked changes by two authors in one control", () => {
  assert.deepEqual(texts(p(sdt("a", ins(r("[1,"), "Anna") + del("[2]", "Ben") + ins(r("3]"), "Ben")))), [["a", "[1,3]", "[2]"]]);
});

test("a tracked move: the control it left is empty, the one it went to has the text, range marks are no text", () => {
  const from = `<w:moveFromRangeStart w:id="3" w:author="A" w:name="move1"/><w:moveFrom w:id="4" w:author="A"><w:r><w:t>[2]</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="3"/>`;
  const to = `<w:moveToRangeStart w:id="5" w:author="A" w:name="move1"/><w:moveTo w:id="6" w:author="A"><w:r><w:t>[2]</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="5"/>`;
  assert.deepEqual(texts(p(sdt("b", from)) + p(sdt("b", to))), [["b", "", "[2]"], ["b", "[2]", ""]]);
});

test("a paragraph moved inside a table cell, its mark moved too", () => {
  const mark = (kind) => `<w:pPr><w:rPr><w:${kind} w:id="9" w:author="A"/></w:rPr></w:pPr>`;
  const cell = (x) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${x}</w:tc></w:tr></w:tbl>`;
  const from = `<w:p>${mark("moveFrom")}<w:moveFromRangeStart w:id="1" w:name="m"/><w:moveFrom w:id="2"><w:r><w:t>Row </w:t></w:r>${sdt("c", "<w:r><w:t>[1]</w:t></w:r>")}</w:moveFrom><w:moveFromRangeEnd w:id="1"/></w:p>`;
  const to = `<w:p>${mark("moveTo")}<w:moveToRangeStart w:id="3" w:name="m"/><w:moveTo w:id="4"><w:r><w:t>Row </w:t></w:r>${sdt("c", "<w:r><w:t>[1]</w:t></w:r>")}</w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>`;
  assert.deepEqual(texts(cell(from + p(r("stays"))) + cell(to + p(r("")))), [["c", "", "[1]"], ["c", "[1]", ""]]);
});

test("a deleted paragraph mark between two citations: both read, a bibliography's paragraphs join", () => {
  assert.deepEqual(texts(pDeletedMark(sdt("a", r("[1]"))) + p(sdt("b", r("[2]")))), [["a", "[1]", "[1]"], ["b", "[2]", "[2]"]]);
  assert.deepEqual(texts(sdt("bib", pDeletedMark(r("1. A")) + p(r("2. B")))), [["bib", "1. A2. B\r", "1. A\r2. B\r"]]);
  // An inserted paragraph mark splits it only once accepted.
  const insertedMark = `<w:p><w:pPr><w:rPr><w:ins w:id="6" w:author="A"/></w:rPr></w:pPr>${r("1. A")}</w:p>`;
  assert.deepEqual(texts(sdt("bib", insertedMark + p(r("2. B")))), [["bib", "1. A\r2. B\r", "1. A2. B\r"]]);
});

test("a control showing its placeholder reads empty both ways", () => {
  assert.deepEqual(texts(p(sdt("e", r("[empty citation]"), "<w:showingPlcHdr/>"))), [["e", "", ""]]);
  assert.deepEqual(texts(p(sdt("f", r("[1]"), '<w:showingPlcHdr w:val="0"/>'))), [["f", "[1]", "[1]"]]);
});

test("a bibliography's paragraphs end in \\r, a tab is \\t, field codes and tab stops are no text", () => {
  const para = (x) => p(x, '<w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs>');
  const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r("today")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  assert.deepEqual(texts(sdt("bib", para(`${r("1.")}<w:r><w:tab/></w:r>${r("A &amp; B")}`) + para(field))), [["bib", "1.\tA & B\rtoday\r", "1.\tA & B\rtoday\r"]]);
});

test("controls in text boxes are left out, as body.contentControls leaves them out", () => {
  const boxed = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${p(sdt("tb", r("[?]")))}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent>${p(sdt("tb", r("[?]")))}</w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  assert.deepEqual(texts(p(`${sdt("a", r("[1]"))}${boxed}${sdt("b", r("[2]"))}`)), [["a", "[1]", "[1]"], ["b", "[2]", "[2]"]]);
});

test("a comment that holds <w:del> is no deletion", () => {
  assert.deepEqual(texts(`<!-- <w:del> -->${p(sdt("x", `<!-- <w:del w:id="1"> --><w:r><w:t>[1]</w:t></w:r><!-- </w:del> -->`))}`), [["x", "[1]", "[1]"]]);
});

test("CDATA in w:t is text", () => {
  assert.deepEqual(texts(p(sdt("x", "<w:r><w:t><![CDATA[[1]<w:del>&amp;]]></w:t></w:r>"))), [["x", "[1]<w:del>&amp;", "[1]<w:del>&amp;"]]);
});

test("numeric and named entities are decoded; an unknown one is an error", () => {
  assert.deepEqual(texts(p(sdt("a&amp;b", r("&#91;1&#x5D; &lt;&gt;&quot;&apos;&amp;")))), [["a&b", "[1] <>\"'&", "[1] <>\"'&"]]);
  unreadable(() => controlTexts(bodyPackage(p(sdt("x", r("[1]&nbsp;"))))));
});

test("another prefix for Word's namespace, on the root and redeclared on an inner element", () => {
  const x = markup("x");
  const q = markup("q");
  const inner = `<q:sdt xmlns:q="${W}"><q:sdtPr><q:tag q:val="inner"/></q:sdtPr><q:sdtContent>${q.del("[2]")}${q.ins(q.r("[3]"))}</q:sdtContent></q:sdt>`;
  const read = controlTexts(pkg(x.document(x.p(x.sdt("outer", x.r("[1]"))) + x.p(inner)))).map((c) => [c.tag, c.current, c.original]);
  assert.deepEqual(read, [["outer", "[1]", "[1]"], ["inner", "[3]", "[2]"]]);
  // The default namespace works too. An unprefixed attribute is in no namespace, so it is no w:val.
  const plain = `<document xmlns="${W}"><body><p><sdt><sdtPr><tag val="d"/></sdtPr><sdtContent><r><t>[4]</t></r></sdtContent></sdt></p></body></document>`;
  assert.deepEqual(controlTexts(pkg(plain)).map((c) => [c.tag, c.current]), [["", "[4]"]]);
  // "w:" bound to another namespace is not Word's.
  const qualified = `<document xmlns="${W}" xmlns:w="${W}"><body><p><sdt><sdtPr><tag w:val="d"/></sdtPr><sdtContent><w:del xmlns:w="urn:other"><r><t>[4]</t></r></w:del></sdtContent></sdt></p></body></document>`;
  assert.deepEqual(controlTexts(pkg(qualified)).map((c) => [c.tag, c.current, c.original]), [["d", "[4]", "[4]"]]);
});

test("only the main document part is read; a large binary part beside it costs no wrong reads", () => {
  const binary = "iVBORw0KGgo" + "A".repeat(8 * 1024 * 1024);
  const header = `<pkg:xmlData><w:hdr xmlns:w="${W}">${p(sdt("in-header", r("[9]")))}</w:hdr></pkg:xmlData>`;
  const xml = pkg(markup().document(p(sdt("a", r("[1]")))), [
    ["/word/media/image1.png", "image/png", `<pkg:binaryData>${binary}</pkg:binaryData>`],
    ["/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", header],
  ]);
  assert.deepEqual(controlTexts(xml).map((c) => [c.tag, c.current]), [["a", "[1]"]]);
});

test("a package without a main document part, or with a foreign root, is unreadable", () => {
  unreadable(() => controlTexts('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>'));
  unreadable(() => controlTexts('<html xmlns="http://www.w3.org/1999/xhtml"/>'));
  unreadable(() => controlTexts(""));
});

test("input cut off anywhere throws, at ten places", () => {
  const full = pkg(markup().document(p(r("Text ") + sdt("siglum:cite:x", del("[1]") + ins(r("[2]")))) + p(r("after"))), [["/word/media/image1.png", "image/png", "<pkg:binaryData>AAAABBBB</pkg:binaryData>"]]);
  controlTexts(full);
  const at = (s, delta = 0) => { const i = full.indexOf(s); assert.ok(i >= 0, s); return i + delta; };
  const cuts = [
    at("<pkg:package", 4), // inside the root's start tag
    at('w:val="siglum:cite:x"', 10), // inside an attribute value
    at("<w:sdtContent>", 14), // right after w:sdtContent opens: the reviewer's case
    at("<w:delText>", 13), // inside deleted text
    at("</w:sdtContent>", 5), // inside an end tag
    at("</w:sdt>", 0), // after w:sdtContent closes
    at("</w:body>", 9), // after the body
    at("</w:document>", 13), // after the document element
    at("<pkg:binaryData>", 20), // inside a binary part
    full.length - 1, // one character short
  ];
  assert.equal(new Set(cuts).size, 10);
  for (const n of cuts) unreadable(() => controlTexts(full.slice(0, n)));
});

test("a paragraph's style, from its own properties only", () => {
  const { p: para } = markup();
  assert.equal(paragraphStyle(bodyPackage(para(r("1. A"), '<w:pStyle w:val="Bibliography"/><w:numPr><w:ilvl w:val="0"/></w:numPr>') + para(r("x"), '<w:pStyle w:val="Other"/>'))), "Bibliography");
  assert.equal(paragraphStyle(bodyPackage(para(r("1. A"), '<w:pPrChange w:id="1"><w:pPr><w:pStyle w:val="Old"/></w:pPr></w:pPrChange>'))), null);
  const x = markup("x");
  assert.equal(paragraphStyle(pkg(x.document(x.p(x.r("A"), '<x:pStyle x:val="Quote"/>')))), "Quote");
  unreadable(() => paragraphStyle(bodyPackage(para(r("1. A"), '<w:pStyle w:val="Bibliography"/>')).slice(0, -20)));
});

test("bookmark names, whatever prefix they use", () => {
  assert.deepEqual(bookmarkNames(bodyPackage(p(r("x")))), []);
  const kept = `<w:bookmarkStart w:id="0" w:name="keepMe"/><w:bookmarkEnd w:id="0"/>`;
  const goback = `<w:bookmarkStart w:id="1" w:name="_GoBack"/><w:bookmarkEnd w:id="1"/>`;
  assert.deepEqual(bookmarkNames(bodyPackage(p(r("x") + kept + goback))), ["keepMe", "_GoBack"]);
  const x = markup("x");
  assert.deepEqual(bookmarkNames(bodyPackage(x.p(x.r("y") + `<x:bookmarkStart x:id="2" x:name="there"/><x:bookmarkEnd x:id="2"/>`), "x")), ["there"]);
});

test("a citation's text rewritten in the package, nothing else touched", () => {
  const before = bodyPackage(p(`before ${sdt("a", r("[1]"))} middle ${sdt("b", r("[2]"))} after`));
  const { xml, applied, skipped } = rewriteCitationTexts(before, new Map([["a", [{ text: "[10]" }]]]));
  assert.deepEqual(applied, ["a"]);
  assert.deepEqual(skipped, []);
  assert.equal(xml, before.replace(`<w:t xml:space="preserve">[1]</w:t>`, `<w:t xml:space="preserve">[10]</w:t>`));
  const read = (x) => controlTexts(x).map((c) => [c.tag, c.current, c.original]);
  assert.equal(read(xml).find(([tag]) => tag === "a")[1], "[10]");
  assert.equal(read(xml).find(([tag]) => tag === "b")[1], "[2]");
});

test("rewritten runs keep formatting, with the document's own prefix", () => {
  const x = markup("x");
  const before = bodyPackage(x.p(x.sdt("a", x.r("[1]"))), "x");
  const { xml, applied, skipped } = rewriteCitationTexts(before, new Map([["a", [{ text: "Au", italic: true }, { text: " [3]" }]]]));
  assert.deepEqual([applied, skipped], [["a"], []]);
  assert.ok(xml.includes(`<x:r><x:rPr><x:i/></x:rPr><x:t xml:space="preserve">Au</x:t></x:r>`));
  assert.ok(xml.includes(`<x:t xml:space="preserve"> [3]</x:t>`));
  assert.ok(!xml.includes(`<w:r><w:t xml:space="preserve">[1]</w:t></w:r>`));
});

test("what the rewrite skips: placeholders, duplicates, missing tags, complex content", () => {
  const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
  const multi = sdt("m", r("[1]") + `<w:r><w:t>more</w:t></w:r>`);
  const nested = sdt("n", r("[1]") + sdt("inner", r("[x]")));
  const before = bodyPackage(p(sdt("ph", r("[empty citation]"), "<w:showingPlcHdr/>") + sdt("d", r("[1]")) + p(sdt("d", r("[2]"))) + sdt("f", field + r("[3]")) + multi + nested));
  const jobs = new Map([["ph", [{ text: "[9]" }]], ["d", [{ text: "[9]" }]], ["gone", [{ text: "[9]" }]], ["f", [{ text: "[9]" }]], ["m", [{ text: "[9]" }]], ["n", [{ text: `[9]` }]]]);
  const { xml, applied, skipped } = rewriteCitationTexts(before, jobs);
  assert.deepEqual(applied, []);
  assert.deepEqual(skipped.map((s) => [s.tag, s.reason]), [["ph", "placeholder"], ["d", "duplicate"], ["gone", "missing"], ["f", "complex"], ["m", "complex"], ["n", "nested"]]);
  assert.equal(xml, before);
});

test("a control off in a text box is rewritten by tag like any other", () => {
  const boxed = `<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wps:txbx xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:txbxContent>${p(sdt("b", r("[?]")))}</w:txbxContent></wps:txbx></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;
  const before = bodyPackage(p(sdt("a", r("[1]")) + boxed));
  const { applied, skipped } = rewriteCitationTexts(before, new Map([["b", [{ text: "[2]" }]]]));
  assert.deepEqual([applied, skipped], [["b"], []]);
});
