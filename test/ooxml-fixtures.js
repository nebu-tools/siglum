// OOXML as Word's getOoxml returns it: a pkg:package with the main document
// part and others beside it. `markup(prefix)` writes WordprocessingML with any
// prefix bound to Word's namespace.

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const NAMESPACES = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

export function markup(w = "w") {
  const m = {
    w,
    sdt: (tag, content, pr = "") => `<${w}:sdt><${w}:sdtPr><${w}:rPr><${w}:rStyle ${w}:val="Hyperlink"/></${w}:rPr><${w}:tag ${w}:val="${tag}"/><${w}:id ${w}:val="1"/>${pr}</${w}:sdtPr><${w}:sdtContent>${content}</${w}:sdtContent></${w}:sdt>`,
    r: (t) => `<${w}:r><${w}:t xml:space="preserve">${t}</${w}:t></${w}:r>`,
    p: (x, pPr = "") => `<${w}:p>${pPr && `<${w}:pPr>${pPr}</${w}:pPr>`}${x}</${w}:p>`,
    del: (t, author = "A") => `<${w}:del ${w}:id="1" ${w}:author="${author}" ${w}:date="2026-01-01T00:00:00Z"><${w}:r><${w}:delText>${t}</${w}:delText></${w}:r></${w}:del>`,
    ins: (x, author = "A") => `<${w}:ins ${w}:id="2" ${w}:author="${author}" ${w}:date="2026-01-01T00:00:00Z">${x}</${w}:ins>`,
    /** A paragraph whose mark is a tracked deletion: it joins the next paragraph once accepted. */
    pDeletedMark: (x) => `<${w}:p><${w}:pPr><${w}:rPr><${w}:del ${w}:id="5" ${w}:author="A" ${w}:date="2026-01-01T00:00:00Z"/></${w}:rPr></${w}:pPr>${x}</${w}:p>`,
    document: (body, rootAttrs = `xmlns:${w}="${W}"`) => `<${w}:document ${rootAttrs} ${NAMESPACES}><${w}:body>${body}<${w}:sectPr><${w}:pgSz ${w}:w="11906" ${w}:h="16838"/></${w}:sectPr></${w}:body></${w}:document>`,
  };
  return m;
}

const part = (name, type, data) => `<pkg:part pkg:name="${name}" pkg:contentType="${type}">${data}</pkg:part>`;

/**
 * A package with `document` (a w:document element) as its main part, the
 * relationships part, a styles part, and `parts` more: [name, type, data].
 */
export function pkg(document, parts = []) {
  return '<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?>'
    + `<pkg:package xmlns:pkg="${PKG}">`
    + part("/_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData>')
    + part("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", `<pkg:xmlData>${document}</pkg:xmlData>`)
    + part("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", `<pkg:xmlData><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles></pkg:xmlData>`)
    + parts.map(([name, type, data]) => part(name, type, data)).join("")
    + "</pkg:package>";
}

/** The package that body.getOoxml returns for a body of `x`. */
export const bodyPackage = (x, w = "w") => pkg(markup(w).document(x));
