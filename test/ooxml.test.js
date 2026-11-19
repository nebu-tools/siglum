import { test } from "node:test";
import assert from "node:assert/strict";
import { controlTexts } from "../src/ooxml.js";

const sdt = (tag, content, pr = "") => `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="1"/>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const r = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const del = (t) => `<w:del w:id="1" w:author="A"><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
const ins = (x) => `<w:ins w:id="2" w:author="A">${x}</w:ins>`;
const body = (x) => `<pkg:package><w:document><w:body>${x}</w:body></w:document></pkg:package>`;
const texts = (x) => controlTexts(body(x)).map((c) => [c.tag, c.current, c.original]);

test("plain controls, in document order, nested ones too", () => {
  assert.deepEqual(texts(`<w:p>${r("One ")}${sdt("a", r("[1]"))}${r(" and ")}${sdt("outer", r("x ") + sdt("b", r("[2]")))}</w:p>`), [["a", "[1]", "[1]"], ["outer", "x [2]", "x [2]"], ["b", "[2]", "[2]"]]);
});

test("a tracked deletion, an insertion and a replacement", () => {
  assert.deepEqual(texts(`<w:p>${sdt("gone", del("[1]"))}${ins(sdt("new", r("[2]")))}${sdt("edit", del("[3]") + ins(r("[4]")))}</w:p>`), [["gone", "", "[1]"], ["new", "[2]", ""], ["edit", "[4]", "[3]"]]);
});

test("a tracked move: the control it left is empty, the one it went to has the text", () => {
  const from = `<w:moveFrom w:id="3"><w:r><w:t>[2]</w:t></w:r></w:moveFrom>`;
  const to = `<w:moveTo w:id="4"><w:r><w:t>[2]</w:t></w:r></w:moveTo>`;
  assert.deepEqual(texts(`<w:p>${sdt("b", from)}</w:p><w:p>${sdt("b", to)}</w:p>`), [["b", "", "[2]"], ["b", "[2]", ""]]);
});

test("a control showing its placeholder reads empty both ways", () => {
  assert.deepEqual(texts(`<w:p>${sdt("e", r("[empty citation]"), "<w:showingPlcHdr/>")}</w:p>`), [["e", "", ""]]);
});

test("a bibliography's paragraphs end in \\r, a tab is \\t, field codes and tab stops are no text", () => {
  const p = (x) => `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs></w:pPr>${x}</w:p>`;
  const field = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r("today")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  assert.deepEqual(texts(sdt("bib", p(`${r("1.")}<w:r><w:tab/></w:r>${r("A &amp; B")}`) + p(field))), [["bib", "1.\tA & B\rtoday\r", "1.\tA & B\rtoday\r"]]);
});

test("controls in text boxes are left out, as body.contentControls leaves them out", () => {
  const boxed = `<w:r><mc:AlternateContent><mc:Choice><w:drawing><wps:txbx><w:txbxContent><w:p>${sdt("tb", r("[?]"))}</w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox><w:txbxContent><w:p>${sdt("tb", r("[?]"))}</w:p></w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  assert.deepEqual(texts(`<w:p>${sdt("a", r("[1]"))}${boxed}${sdt("b", r("[2]"))}</w:p>`), [["a", "[1]", "[1]"], ["b", "[2]", "[2]"]]);
});
