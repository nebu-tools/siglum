// Legacy upgrade: brings a document written before the stored format was
// renamed up to the current format. This is the only file that knows the old
// names. word.js calls upgradeLegacy() once, before it reads the store, so
// every other module sees the current format only.
//
// The old format:
//
//   store          <zitate xmlns="urn:newga:zitate:1">{"quellen": {…}, "zitate": {…}}</zitate>
//   citation       content control tagged `zitat:<id>`, titled `Zitat: key1; key2`
//   bibliography   content control tagged `literaturverzeichnis`
//   typed marker   `[@key]`, `[@a; @b]`, `[@key, S. 12]`

import { BIB_TAG, CITE_PREFIX, TITLE_PREFIX, xmlPayload } from "./model.js";

const OLD = {
  ns: "urn:newga:zitate:1",
  root: "zitate",
  citePrefix: "zitat:",
  bibTag: "literaturverzeichnis",
  titlePrefix: "Zitat: ",
  // The old version always wrote German citations and bibliographies.
  locale: "de-DE",
};

/** The store in an old XML part, in the current shape, or null when the part is empty. */
export function storeFromLegacyXml(xml) {
  const json = xmlPayload(xml, OLD.root);
  if (!json) return null;
  const d = JSON.parse(json);
  return { sources: d.quellen ?? {}, citations: d.zitate ?? {}, settings: { locale: OLD.locale } };
}

/** The current tag for an old content control tag, or null for any other tag. */
export function upgradeTag(tag) {
  if (tag === OLD.bibTag) return BIB_TAG;
  if (tag?.startsWith(OLD.citePrefix)) return CITE_PREFIX + tag.slice(OLD.citePrefix.length);
  return null;
}

/** An old citation title with the current prefix. Other titles come back unchanged. */
export function upgradeTitle(title) {
  return title?.startsWith(OLD.titlePrefix) ? TITLE_PREFIX + title.slice(OLD.titlePrefix.length) : title;
}

const OLD_MARKER = /\[\s*@[^\[\]\r\n]*\]/g;
const OLD_ITEM = /^\s*@([A-Za-z0-9_:.\-]+)\s*(?:,\s*(.+?))?\s*$/;

/**
 * Rewrites old markers in a text to the current syntax: `[@a; @b, S. 12]`
 * becomes `@a @b[S. 12]`. Brackets that are not an old marker stay as they are.
 */
export function upgradeMarkers(text) {
  return text.replace(OLD_MARKER, (marker) => {
    const items = marker.slice(1, -1).split(";").map((part) => OLD_ITEM.exec(part));
    // A key with a trailing "." or ":" has no spelling in the current syntax.
    if (items.some((m) => !m || /[.:]$/.test(m[1]) || /[\[\]]/.test(m[2] ?? ""))) return marker;
    return items.map((m) => (m[2] ? `@${m[1]}[${m[2]}]` : `@${m[1]}`)).join(" ");
  });
}

/**
 * Upgrades the open document in place. `adopt(store)` must save an old store
 * in the current format; the old part goes only after that. `untracked(fn)`
 * must run `fn` with track changes off, because the rewrite of typed markers
 * is no edit for review. A document in the current format costs one sync.
 */
export async function upgradeLegacy(context, { adopt, untracked }) {
  const body = context.document.body;
  const parts = context.document.customXmlParts.getByNamespace(OLD.ns);
  parts.load("items");
  const controls = body.contentControls;
  controls.load("items/tag,items/title");
  const probe = body.search("[@", { matchCase: true });
  probe.load("items");
  await context.sync();

  if (parts.items.length) {
    const xmls = parts.items.map((p) => p.getXml());
    await context.sync();
    for (const xml of xmls) {
      const store = storeFromLegacyXml(xml.value);
      if (store) await adopt(store);
    }
    for (const p of parts.items) p.delete();
    await context.sync();
  }

  let retagged = false;
  for (const cc of controls.items) {
    const tag = upgradeTag(cc.tag);
    if (!tag) continue;
    cc.tag = tag;
    cc.title = upgradeTitle(cc.title);
    retagged = true;
  }
  if (retagged) await context.sync();

  if (!probe.items.length) return;
  // Word's wildcards: "\" makes "[", "@" and "]" literal, "*" is any text.
  const found = body.search("\\[\\@*\\]", { matchWildcards: true });
  found.load("items/text");
  await context.sync();
  const rewrites = found.items.map((r) => [r, upgradeMarkers(r.text)]).filter(([r, text]) => text !== r.text);
  if (!rewrites.length) return;
  await untracked(async () => {
    for (const [r, text] of rewrites) r.insertText(text, "Replace");
    await context.sync();
  });
}
