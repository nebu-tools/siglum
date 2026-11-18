// What the binding experiment's generator (Node) and page (the pane) share.
export const NS = "urn:siglum:exp:cites";
export const ITEM_ID = "{5E1A9C3B-0000-4000-8000-0000000000B1}";
export const PREFIX = `xmlns:c='${NS}'`;
export const TAG = "exp:cite:";
export const BIB = "exp:bib";
export const PAGE = "bind.html";

export const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** The xpath of a citation's element: by its id attribute, or by its position. */
export const xpathFor = (id, i, mode = "id") => (mode === "pos" ? `/c:cites[1]/c:c[${i + 1}]` : `/c:cites[1]/c:c[@id='${id}']`);
/** The store part: one element per citation, in document order. */
export const storeXml = (entries) => `<c:cites xmlns:c="${NS}">${[...entries].map(([id, v]) => `<c:c id="${esc(id)}">${esc(v)}</c:c>`).join("")}</c:cites>`;

export function runXml(run, extra = "") {
  const pr = [run.bold && "<w:b/>", run.italic && "<w:i/>", run.superscript && '<w:vertAlign w:val="superscript"/>'].filter(Boolean).join("") + extra;
  const rpr = pr ? `<w:rPr>${pr}</w:rPr>` : "";
  return `<w:r>${rpr}<w:t xml:space="preserve">${esc(run.text)}</w:t></w:r>`;
}

/** The flat-OPC package that a rich-text bound control's element holds, for one paragraph of runs. */
export function richPackage(runsXml) {
  return `<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${runsXml}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
}
