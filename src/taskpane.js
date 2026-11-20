// Task pane: search the document's sources, cite, keep the bibliographies current.
/* global Office, Word */

import style from "../csl/nlm-citation-sequence-brackets.csl";
import deDE from "../csl/locales-de-DE.xml";
import enUS from "../csl/locales-en-US.xml";
import { createRenderer } from "./render.js";
import { parseSupplement, LABELS } from "./model.js";
import { parseLibrary, mergeItems, year } from "./library.js";
import { buildIndex, search } from "./search.js";
import { lookup } from "./lookup.js";
import { LANGUAGES, pickLanguage, translator, UserError } from "./messages.js";
import * as doc from "./word.js";
import { createQueue } from "./queue.js";

// A style picker would replace `style` here. Nothing else knows the style.
const session = { renderer: createRenderer({ style, locales: { "en-US": enUS, "de-DE": deDE } }), locale: "en-US" };
const LANGUAGE_KEY = "siglum.language";
const $ = (id) => document.getElementById(id);

const state = { library: {}, index: [], hits: [], active: 0, marked: [], note: null };
let t = translator("en");

// ---- language -----------------------------------------------------------------

function savedLanguage() {
  try { return localStorage.getItem(LANGUAGE_KEY); } catch { return null; }
}

/** Shows the pane in `language`. Elements name their message keys in data-i18n attributes. */
function setLanguage(language) {
  const lang = pickLanguage(language);
  t = translator(lang);
  document.documentElement.lang = lang;
  for (const el of document.querySelectorAll("[data-i18n]")) el.textContent = t(el.dataset.i18n);
  for (const attr of ["placeholder", "aria-label", "title"]) {
    const data = "data-i18n-" + attr;
    for (const el of document.querySelectorAll(`[${data}]`)) el.setAttribute(attr, t(el.getAttribute(data)));
  }
  $("language").value = lang;
  for (const o of $("locale").options) o.textContent = localeName(o.value);
  showMarked();
  showResults();
  showStatus();
  if (editing) {
    // Keep what was typed: re-rendering rebuilds the fields from the items.
    editing = { id: editing.id, items: readEditor(editing.items) };
    renderEditor();
  }
}

function localeName(locale) {
  const name = t("locale." + locale);
  return name === "locale." + locale ? locale : name;
}

function fillSettings() {
  $("language").replaceChildren(...Object.entries(LANGUAGES).map(([value, name]) => new Option(name, value)));
  $("locale").replaceChildren(...session.renderer.locales.map((l) => new Option(localeName(l), l)));
}

// ---- search ---------------------------------------------------------------------

function setStore(store) {
  state.library = store.sources;
  state.index = buildIndex(state.library);
  $("locale").value = store.settings.locale;
  showResults();
}

function names(item) {
  const a = item.author ?? item.editor ?? [];
  const n = a.slice(0, 2).map((x) => x.literal ?? [x.family, (x.given ?? "").replace(/[^\p{Lu}]/gu, "")].filter(Boolean).join(" "));
  return n.join(", ") + (a.length > 2 ? " et al." : "") || t("results.noAuthor");
}

function showResults() {
  state.hits = search(state.index, $("q").value, 60);
  state.active = Math.min(state.active, Math.max(0, state.hits.length - 1));
  const ul = $("results");
  ul.replaceChildren();
  if (!state.hits.length) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = t(Object.keys(state.library).length ? "results.none" : "results.emptyLibrary");
    ul.append(li);
    return;
  }
  state.hits.forEach((item, i) => {
    const li = document.createElement("li");
    li.id = "r" + i;
    li.setAttribute("role", "option");
    li.setAttribute("aria-selected", String(i === state.active));
    if (state.marked.includes(item.id)) li.classList.add("on");
    const who = document.createElement("span");
    who.className = "who";
    who.textContent = `${names(item)} (${year(item) || t("results.noYear")})`;
    const key = document.createElement("span");
    key.className = "key";
    key.textContent = item.id;
    const what = document.createElement("span");
    what.className = "what";
    what.textContent = [String(item.title ?? "").replace(/<[^>]+>/g, ""), item["container-title"]].filter(Boolean).join(" · ");
    li.append(who, key, what);
    li.addEventListener("click", (e) => (e.ctrlKey || e.metaKey || e.shiftKey ? toggleMark(item.id) : cite([item.id])));
    ul.append(li);
  });
  $("q").setAttribute("aria-activedescendant", "r" + state.active);
  $("r" + state.active)?.scrollIntoView({ block: "nearest" });
}

function showMarked() {
  const box = $("marked");
  box.hidden = !state.marked.length;
  box.replaceChildren(...state.marked.map((k) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = k;
    const x = document.createElement("button");
    x.type = "button";
    x.textContent = "×";
    x.setAttribute("aria-label", t("marked.remove", { key: k }));
    x.addEventListener("click", () => toggleMark(k));
    chip.append(x);
    return chip;
  }));
}

function toggleMark(key) {
  state.marked = state.marked.includes(key) ? state.marked.filter((k) => k !== key) : [...state.marked, key];
  showMarked();
  showResults();
  $("q").focus();
}

// ---- edit citation ------------------------------------------------------------

// The citation the editor shows, {id, items}, or null when it is hidden.
let editing = null;

function hideEditor() {
  editing = null;
  $("editCitation").hidden = true;
}

function editRow(item) {
  const row = document.createElement("div");
  row.className = "edit-item";
  row.dataset.key = item.id;
  const key = document.createElement("span");
  key.className = "edit-key";
  key.textContent = item.id;
  row.append(key);
  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "edit-remove-item";
  remove.textContent = "×";
  remove.setAttribute("aria-label", t("marked.remove", { key: item.id }));
  remove.addEventListener("click", () => row.remove());
  row.append(remove);
  const locator = document.createElement("input");
  locator.className = "edit-locator";
  locator.value = item.locator ?? "";
  locator.setAttribute("aria-label", t("edit.locator"));
  locator.placeholder = t("edit.locator");
  row.append(locator);
  const label = document.createElement("select");
  label.className = "edit-label";
  label.setAttribute("aria-label", t("edit.label"));
  // A label from another source stays selectable, so saving never drops it.
  const labels = item.label && !Object.hasOwn(LABELS, item.label) ? [item.label, ...Object.keys(LABELS)] : Object.keys(LABELS);
  label.replaceChildren(new Option("", ""), ...labels.map((l) => new Option(l, l)));
  label.value = item.label ?? (item.locator ? "page" : "");
  row.append(label);
  const prefix = document.createElement("input");
  prefix.className = "edit-prefix";
  prefix.value = item.prefix ?? "";
  prefix.setAttribute("aria-label", t("edit.prefix"));
  prefix.placeholder = t("edit.prefix");
  row.append(prefix);
  const suffix = document.createElement("input");
  suffix.className = "edit-suffix";
  suffix.value = item.suffix ?? "";
  suffix.setAttribute("aria-label", t("edit.suffix"));
  suffix.placeholder = t("edit.suffix");
  row.append(suffix);
  return row;
}

function renderEditor() {
  $("editItems").replaceChildren(...editing.items.map(editRow));
}

/** The items as edited: locators with labels, affixes; empty fields are dropped.
 * Fields with no editor control travel with their item, so neither saving
 * nor a language switch while editing drops them. */
function readEditor(before = []) {
  const prev = new Map(before.map((x) => [x.id, x]));
  return [...$("editItems").children].map((row) => {
    const item = { id: row.dataset.key };
    const locator = row.querySelector(".edit-locator").value.trim();
    const label = row.querySelector(".edit-label").value;
    const prefix = row.querySelector(".edit-prefix").value.trim();
    const suffix = row.querySelector(".edit-suffix").value.trim();
    if (locator) { item.locator = locator; if (label) item.label = label; }
    if (prefix) item.prefix = prefix;
    if (suffix) item.suffix = suffix;
    if (prev.get(item.id)?.suppressAuthor) item.suppressAuthor = true;
    return item;
  });
}

async function updateEditor() {
  if (document.body.hasAttribute("aria-busy")) return;
  let found = null;
  try {
    await Word.run(async (c) => { found = await doc.citationAtCursor(c, session); });
  } catch {
    found = null;
  }
  if (!found) { hideEditor(); return; }
  editing = found;
  renderEditor();
  $("editCitation").hidden = false;
}

// Showing the editor follows the cursor, on its own debounce, apart from the
// watcher's refresh: selecting text never writes by itself.
function watchEditor() {
  let timer = null;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    clearTimeout(timer);
    timer = setTimeout(() => updateEditor().catch((e) => console.error(e)), 400);
  });
}

// ---- status ---------------------------------------------------------------------

/**
 * The status line: an error, what the last operation has to say (`note`), and
 * then where the document stands: the progress of a renumbering, citation
 * numbers that are not up to date, or the summary of the last refresh. Texts
 * are functions, so the status follows a change of language.
 */
function note(text, error = false) {
  state.note = { text, error };
  showStatus();
}

function showStatus() {
  const q = queue.state;
  const lines = [];
  if (q.error) lines.push(errorText(q.error));
  if (state.note) lines.push(state.note.text());
  let problem = Boolean(q.error || state.note?.error);
  if (q.progress) lines.push(t("progress.writing", q.progress));
  else if (q.stale) { lines.push(t("status.stale")); problem = true; }
  else if (q.result && !state.note && !q.error) { lines.push(summary(q.result)); problem ||= hasProblems(q.result); }
  $("status").textContent = lines.join("\n");
  $("status").classList.toggle("error", problem);
}

const errorText = (e) => (e instanceof UserError ? t(e.key, e.params) : e?.message ?? String(e));
const sourceCount = (library) => t("summary.sources", { n: Object.keys(library).length });

function summary(r) {
  const parts = [sourceCount(r.store.sources), t("summary.citations", { n: r.cited })];
  if (r.sections > 1) parts.push(t("summary.sections", { n: r.sections }));
  if (r.pending) parts.push(t("summary.pending", { n: r.pending }));
  let msg = parts.join(" · ");
  if (r.converted) msg = t("summary.converted", { n: r.converted }) + " " + msg;
  const unknown = [...(r.unknown ?? []).map((u) => u.keys.join(", ") || u.text), ...(r.unknownMarkers ?? []).map((u) => u.keys.join(", "))];
  if (unknown.length) msg += "\n" + t("summary.unknown", { keys: [...new Set(unknown)].join("; ") });
  if (r.lostMarkers?.length) msg += "\n" + t("summary.lost", { markers: [...new Set(r.lostMarkers)].join("; ") });
  return msg;
}

const hasProblems = (r) => Boolean(r.unknown?.length || r.unknownMarkers?.length || r.lostMarkers?.length);

// ---- document operations ----------------------------------------------------------

const queue = createQueue({
  word: (fn) => Word.run(fn),
  refresh: (context) => doc.refresh(context, session),
  writeChunk: doc.writeChunk,
  signature: doc.signature,
  chunk: doc.CHUNK,
  now: () => performance.now(),
  emit(type, detail) {
    if (type === "change") {
      if (detail.busy) document.body.setAttribute("aria-busy", "true");
      else document.body.removeAttribute("aria-busy");
      showStatus();
    } else if (type === "done") {
      if (detail?.store) setStore(detail.store);
      document.dispatchEvent(new CustomEvent("siglum:done", { detail: detail?.error ? { error: errorText(detail.error) } : detail }));
    } else {
      document.dispatchEvent(new CustomEvent("siglum:settled"));
    }
  },
});

/** Runs a document operation through the queue. Its note, if any, replaces the last one's. */
function op(fn) {
  state.note = null;
  return queue.run(fn);
}

async function cite(keys) {
  const list = keys.length ? keys : state.marked;
  if (!list.length) return;
  // The supplement belongs to one source, so it is used for a single one only.
  const supplement = list.length === 1 ? parseSupplement($("supplement").value) : {};
  const items = list.map((id) => ({ id, ...supplement }));
  const r = await op((c) => doc.insertCitation(c, session, items, { onVisible: () => document.dispatchEvent(new CustomEvent("siglum:visible")) }));
  if (r) {
    state.marked = [];
    $("supplement").value = "";
    showMarked();
    showResults();
  }
}

async function importFiles(files) {
  let incoming = [];
  for (const f of files) {
    try {
      incoming = incoming.concat(parseLibrary(await f.text(), f.name));
    } catch (e) {
      note(() => t("import.failed", { file: f.name, reason: errorText(e) }), true);
      return;
    }
  }
  const r = await op((c) => doc.addSources(c, session, (lib) => mergeItems(lib, incoming)));
  if (r) note(() => `${t("import.done", { added: r.added, updated: r.updated })} · ${sourceCount(r.library)}`);
}

async function fetchIdent(text) {
  let item;
  try {
    item = await lookup(text);
  } catch (e) {
    note(() => errorText(e), true);
    return;
  }
  const r = await op((c) => doc.addSources(c, session, (lib) => mergeItems(lib, [item])));
  if (!r) return;
  const same = (a, b) => a && b && String(a).toLowerCase() === String(b).toLowerCase();
  const key = Object.values(r.library).find((x) => same(x.DOI, item.DOI) || same(x.PMID, item.PMID))?.id;
  note(() => t(r.added ? "lookup.added" : "lookup.known", { key }));
  $("q").value = key ?? "";
  $("ident").value = "";
  showResults();
  $("q").focus();
}

function wire() {
  $("q").addEventListener("input", () => { state.active = 0; showResults(); });
  $("q").addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      state.active = Math.max(0, Math.min(state.hits.length - 1, state.active + (e.key === "ArrowDown" ? 1 : -1)));
      showResults();
    } else if (e.key === "Enter") {
      e.preventDefault();
      const hit = state.hits[state.active];
      if ((e.ctrlKey || e.metaKey) && hit) toggleMark(hit.id);
      else if (state.marked.length) cite([...state.marked, ...(hit && !state.marked.includes(hit.id) && $("q").value.trim() ? [hit.id] : [])]);
      else if (hit) cite([hit.id]);
    } else if (e.key === "Escape") {
      $("q").value = "";
      showResults();
    }
  });
  $("supplement").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); const h = state.hits[state.active]; if (h) cite([h.id]); } });
  $("bib").addEventListener("click", () => op((c) => doc.insertBibliography(c, session)));
  $("refresh").addEventListener("click", () => op((c) => doc.convertMarkers(c, session)));
  $("editSave").addEventListener("click", async () => {
    if (!editing) return;
    const id = editing.id;
    const items = readEditor(editing.items);
    const r = await op((c) => doc.editCitation(c, session, id, items));
    if (!r) return;
    if (!items.length) { hideEditor(); return; }
    editing = { id, items };
    renderEditor();
  });
  $("editRemove").addEventListener("click", async () => {
    if (!editing) return;
    const r = await op((c) => doc.editCitation(c, session, editing.id, []));
    if (!r) return;
    hideEditor();
  });
  $("file").addEventListener("change", (e) => { importFiles([...e.target.files]); e.target.value = ""; });
  $("lookup").addEventListener("submit", (e) => { e.preventDefault(); fetchIdent($("ident").value); });
  $("language").addEventListener("change", (e) => {
    try { localStorage.setItem(LANGUAGE_KEY, e.target.value); } catch { /* The choice then lasts for this session. */ }
    setLanguage(e.target.value);
  });
  $("locale").addEventListener("change", (e) => op((c) => doc.setLocale(c, session, e.target.value)));
}

// Deleting or pasting a citation changes the order of tags. A quiet check after
// the cursor moves keeps numbers and bibliographies current without a click.
// It does not run while citation numbers are known not to be up to date: the
// status line then asks for a click on Refresh.
function watch() {
  let timer = null;
  const check = () => queue.check().catch((e) => console.error(e));
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    clearTimeout(timer);
    timer = setTimeout(check, 700);
  });
}

Office.onReady(async () => {
  // The pane follows Office's display language. A new document's citations
  // follow its editing language.
  session.locale = Office.context.contentLanguage ?? Office.context.displayLanguage ?? "en-US";
  fillSettings();
  setLanguage(savedLanguage() ?? Office.context.displayLanguage);
  wire();
  await op((c) => doc.convertMarkers(c, session));
  watch();
  watchEditor();
  $("q").focus();
  document.dispatchEvent(new CustomEvent("siglum:ready"));
});
