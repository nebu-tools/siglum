// The one place where the rig learns the add-in's stored names. Everything the
// generator and the measuring panes write or look for (tags, the custom XML
// namespace, the store's keys, the citation title, the typed marker) comes from
// src/model.js through here, so a rename in the add-in needs no change in the
// rig. If an export itself is renamed, fix it in this file only.
//
// Runs in Node and, bundled, in the task pane.
import * as model from "../src/model.js";

const pick = (...names) => {
  for (const n of names) if (n in model) return model[n];
  throw new Error(`src/model.js exports none of ${names.join(", ")}: update e2e/names.mjs`);
};

/** Tag prefix of a citation control, tag of a reference list, custom XML namespace. */
export const CITE_PREFIX = pick("CITE_PREFIX", "CITATION_PREFIX", "CITE_TAG_PREFIX");
export const BIB_TAG = pick("BIB_TAG", "BIBLIOGRAPHY_TAG");
export const NS = pick("NS", "NAMESPACE");
/** Text of a citation before it is numbered. */
export const CITE_NEW = pick("CITE_NEW", "CITATION_NEW");

export const storeToXml = pick("storeToXml");
export const storeFromXml = pick("storeFromXml");
export const citeTitle = pick("citeTitle", "citationTitle");
export const keysFromTitle = pick("keysFromTitle");
export const newId = pick("newId");
export const parseMarker = pick("parseMarker");
export const plan = pick("plan");
export const emptyStore = pick("emptyStore");

// The store is {library, citations} under whatever names emptyStore() uses.
// The library is the key that looks like one; the other one holds citations.
const storeKeys = Object.keys(emptyStore());
const libRe = /quell|source|lib|ref|work|item/i;
export const LIBRARY_KEY = storeKeys.find((k) => libRe.test(k)) ?? storeKeys[0];
export const CITATIONS_KEY = storeKeys.find((k) => k !== LIBRARY_KEY);

/** A store object with the add-in's key names. */
export function makeStore(library, citations) {
  return { ...emptyStore(), [LIBRARY_KEY]: library, [CITATIONS_KEY]: citations };
}
export const libraryOf = (store) => store[LIBRARY_KEY];
export const citationsOf = (store) => store[CITATIONS_KEY];

/**
 * The typed marker for `keys`, in whatever syntax parseMarker accepts: the
 * first candidate that parses back to exactly these keys.
 */
export function marker(keys) {
  const ats = keys.map((k) => "@" + k);
  const candidates = [
    `[${ats.join("; ")}]`, `[${keys.join("; ")}]`, `[${keys.join(", ")}]`, `{${ats.join("; ")}}`,
    `[[${keys.join("; ")}]]`, `{{${keys.join("; ")}}}`, `\\cite{${keys.join(",")}}`, ats.join("; "),
  ];
  for (const c of candidates) {
    const items = parseMarker(c);
    if (items && items.length === keys.length && items.every((x, i) => x.id === keys[i])) return c;
  }
  return null;
}

/** A citation control's tag and title for one stored citation. */
export const citeTag = (id) => CITE_PREFIX + id;
export const isCiteTag = (tag) => typeof tag === "string" && tag.startsWith(CITE_PREFIX);
