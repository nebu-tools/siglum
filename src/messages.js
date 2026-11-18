// Every text a person reads, per language. English is the reference and the
// fallback. A value is a string, or [singular, plural] chosen by the
// parameter `n`. `{name}` is replaced by a parameter.

/** The interface languages and their own names. */
export const LANGUAGES = { en: "English", de: "Deutsch" };

export const MESSAGES = {
  en: {
    "search.placeholder": "Author, year, title, key, DOI …",
    "search.label": "Search sources",
    "supplement.placeholder": "p. 12",
    "supplement.label": "Page or other place in the source (optional)",
    "supplement.title": "Page, chapter or a note, for a single source",
    "results.label": "Results",
    "results.none": "No results.",
    "results.emptyLibrary": "This document has no sources yet. Import a file or fetch a DOI under “Add sources”.",
    "results.noAuthor": "Anon.",
    "results.noYear": "n.d.",
    "marked.remove": "Remove {key}",
    hint: "Enter cites the highlighted result. Ctrl+Enter marks it, to cite several sources together. You can also type @key in the text and click “Refresh”.",
    "bibliography.insert": "Insert bibliography",
    refresh: "Refresh",
    "sources.add": "Add sources",
    "sources.file": "Import a file (.bib, .ris, .json)",
    "lookup.placeholder": "DOI or PMID",
    "lookup.submit": "Fetch",
    "lookup.added": "Added: {key}",
    "lookup.known": "Already there: {key}",
    settings: "Settings",
    "settings.language": "Language of this pane",
    "settings.locale": "Language of citations and bibliographies",
    "locale.en-US": "English",
    "locale.de-DE": "German",
    "summary.sources": ["{n} source", "{n} sources"],
    "summary.citations": ["{n} citation", "{n} citations"],
    "summary.sections": "{n} sections",
    "summary.pending": "{n} deleted, waiting for review",
    "summary.converted": ["{n} marker converted.", "{n} markers converted."],
    "summary.unknown": "Unknown: {keys}",
    "summary.lost": "Not converted: {markers}",
    "progress.writing": "Updating citations: {done} of {total}",
    "import.done": "{added} new, {updated} updated",
    "import.failed": "Could not read {file}: {reason}",
    "error.noIdentifier": "Neither a DOI nor a PMID: {text}",
    "error.pmidNotFound": "PMID not found ({reason})",
    "error.doiNotFound": "DOI not found ({reason})",
    "error.inBibliography": "The cursor is in the bibliography.",
    "error.inControl": "The cursor is in a citation or a bibliography.",
    // Written into the document, in the language of its citations.
    "document.bibEmpty": "No sources cited.",
    "document.citePlaceholder": "[empty citation]",
    "document.bibTitle": "Bibliography",
  },
  de: {
    "search.placeholder": "Autor, Jahr, Titel, Schlüssel, DOI …",
    "search.label": "Quellen durchsuchen",
    "supplement.placeholder": "S. 12",
    "supplement.label": "Seite oder andere Stelle in der Quelle (optional)",
    "supplement.title": "Seite, Kapitel oder eine Anmerkung, nur bei einer einzelnen Quelle",
    "results.label": "Treffer",
    "results.none": "Keine Treffer.",
    "results.emptyLibrary": "Dieses Dokument hat noch keine Quellen. Unter „Quellen hinzufügen“ eine Datei importieren oder eine DOI holen.",
    "results.noAuthor": "o. A.",
    "results.noYear": "o. J.",
    "marked.remove": "{key} entfernen",
    hint: "Enter zitiert den markierten Treffer. Strg+Enter merkt ihn vor, um mehrere Quellen zusammen zu zitieren. Im Text funktioniert auch @schluessel mit „Aktualisieren“.",
    "bibliography.insert": "Literaturverzeichnis einfügen",
    refresh: "Aktualisieren",
    "sources.add": "Quellen hinzufügen",
    "sources.file": "Datei importieren (.bib, .ris, .json)",
    "lookup.placeholder": "DOI oder PMID",
    "lookup.submit": "Holen",
    "lookup.added": "Hinzugefügt: {key}",
    "lookup.known": "Schon vorhanden: {key}",
    settings: "Einstellungen",
    "settings.language": "Sprache dieses Bereichs",
    "settings.locale": "Sprache der Zitate und Literaturverzeichnisse",
    "locale.en-US": "Englisch",
    "locale.de-DE": "Deutsch",
    "summary.sources": ["{n} Quelle", "{n} Quellen"],
    "summary.citations": ["{n} Zitat", "{n} Zitate"],
    "summary.sections": "{n} Abschnitte",
    "summary.pending": "{n} gelöscht, offen in der Überarbeitung",
    "summary.converted": ["{n} Marker umgewandelt.", "{n} Marker umgewandelt."],
    "summary.unknown": "Unbekannt: {keys}",
    "summary.lost": "Nicht umgewandelt: {markers}",
    "progress.writing": "Zitate werden aktualisiert: {done} von {total}",
    "import.done": "{added} neu, {updated} aktualisiert",
    "import.failed": "{file} ließ sich nicht lesen: {reason}",
    "error.noIdentifier": "Weder DOI noch PMID: {text}",
    "error.pmidNotFound": "PMID nicht gefunden ({reason})",
    "error.doiNotFound": "DOI nicht gefunden ({reason})",
    "error.inBibliography": "Der Cursor steht im Literaturverzeichnis.",
    "error.inControl": "Der Cursor steht in einem Zitat oder Literaturverzeichnis.",
    "document.bibEmpty": "Keine Quellen zitiert.",
    "document.citePlaceholder": "[leeres Zitat]",
    "document.bibTitle": "Literaturverzeichnis",
  },
};

/** The interface language for a language tag such as "de-AT". English when there is no translation. */
export function pickLanguage(tag) {
  const lang = String(tag ?? "").toLowerCase().split(/[-_]/)[0];
  return lang in MESSAGES ? lang : "en";
}

/** Returns t(key, params) for one language. */
export function translator(tag) {
  const table = MESSAGES[pickLanguage(tag)];
  return (key, params = {}) => {
    let msg = table[key] ?? MESSAGES.en[key];
    if (msg == null) return key;
    if (Array.isArray(msg)) msg = msg[params.n === 1 ? 0 : 1];
    return msg.replace(/\{(\w+)\}/g, (all, name) => (name in params ? String(params[name]) : all));
  };
}

/** The texts that are written into a document, in the language of its citations. */
export function documentText(locale) {
  const t = translator(locale);
  return { bibEmpty: t("document.bibEmpty"), citePlaceholder: t("document.citePlaceholder"), bibTitle: t("document.bibTitle") };
}

/**
 * An error whose text is for the person at the pane. It carries the message
 * key, so the pane shows it in its language. `message` is the English text.
 */
export class UserError extends Error {
  constructor(key, params = {}) {
    super(translator("en")(key, params));
    this.name = "UserError";
    this.key = key;
    this.params = params;
  }
}
