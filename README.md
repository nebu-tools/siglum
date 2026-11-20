<img align="right" src="assets/branding/siglum-logo.webp" alt="Siglum logo: a squirrel with acorns numbered one, two and three" width="160" height="160">

# Siglum

Siglum is a citation manager that runs inside Microsoft Word as an add-in. It keeps your sources in the document and keeps the citation numbers and the bibliography current.

It is for people who write long documents in Word and want citations without installing desktop software.

<br clear="right">

## How to cite

Open the Siglum pane from the Home tab. Add your sources first: import a `.bib`, `.ris` or CSL-JSON file, or fetch one source by its DOI or PMID.

You can cite in two ways.

Search in the pane. Type part of an author, year, title or key and press Enter. Siglum puts the citation at the cursor. Press Ctrl+Enter to mark a result when you want to cite several sources together. The small field next to the search box takes a page or another place in the source, such as `p. 12`.

Or type a marker in the text and click Refresh:

| You type | Result |
| --- | --- |
| `@smith2020` | A citation of the source with the key `smith2020` |
| `@smith2020[p. 12]` | The same citation with page 12 |
| `@smith2020 @lee2010` | One citation of both sources |

The rules for markers:

- A marker starts with `@`. The `@` must be at the start of a paragraph or come after a space, a tab, a no-break space, `(`, `[`, `{`, a quotation mark, an en dash, an em dash, `,` or `;`. After anything else it is ordinary text, so e-mail addresses and addresses such as `example.org/@name` stay as they are.
- A key can contain letters, combining accents, digits, `_`, `-`, `:` and `.`. It does not start with a combining accent and does not end in `.` or `:`, so the full stop in `see @smith2020.` stays in the sentence. Siglum compares keys in Unicode's composed form, so `é` typed as one character or as `e` with a combining accent is the same key.
- Markers with only spaces or no-break spaces between them become one citation. A tab or a line break between them makes two citations.
- The text in square brackets right after the key is the supplement. Brackets inside it count in pairs, so `@a[see [2]]` has the supplement `see [2]`. A bracket that does not close in the same paragraph, as in `@a[p. 12`, belongs to the text, and `@a` is a citation without a supplement. A supplement holds no tab or line break.
- Siglum reads `p.`, `pp.`, `ch.`, `sec.`, a bare number and the German forms such as `S.` and `Kap.` as a place in the source. It adds any other supplement after the citation.
- A marker whose key is not in the document's sources stays in the text, and the pane lists it under "Unknown". When one key of a group is unknown, the whole group stays.

Click "Insert bibliography" to put a bibliography at the cursor. A document can have more than one. Each bibliography lists the sources cited since the bibliography before it, and the numbers start again at 1 after it.

A locator prints after the number: `@smith2020[p. 12]` cites as `[1, p. 12]`. Put the cursor in a citation to edit it in the pane: change the page, add a prefix or a note, remove one source from a shared citation, or remove the citation. The pane saves through the same queue as everything else, so an edit waits for a renumbering that is still running.

Siglum renumbers the citations when their order changes. With track changes on, Word records a new citation as an insertion. Siglum writes the changed numbers and the bibliography without tracking, because they follow from the citations.

Siglum writes only what changed. When a new citation changes many numbers, as a new first citation in a long document does, the new citation gets its number at once, and the pane shows "Updating citations: … of …" while the rest follows in batches. You can go on writing and citing meanwhile. When you do something in the pane before the batches are done, Siglum does it first and then works out the numbers again from the document, and the progress line stays until they are all written. If a batch fails, Siglum also works the numbers out again. If that fails too, the pane says that the citation numbers are not up to date, until a refresh succeeds. If you close the pane or the document before it is done, the next refresh puts the numbers right. The pane refreshes when it opens.

## Install

An Office add-in has two parts: a web page, and a small file called `manifest.xml` that tells Word where the page is. First put the page on a web server (see the next section). Then give Word the manifest. This is called sideloading.

### Windows

1. Put `manifest.xml` in a shared folder, for example `\\my-computer\addins`. A folder on your own computer works when you share it.
2. In Word, open File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.
3. Enter the address of the shared folder as the catalog URL, click "Add catalog", select "Show in Menu" and click OK.
4. Start Word again.
5. Open Home > Add-ins > Advanced. On the "Shared Folder" tab, select Siglum and click Add.

### Mac

1. Open the folder `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`. Create the `wef` folder if it does not exist.
2. Copy `manifest.xml` into this folder.
3. Start Word again.
4. Open Home > Add-ins. Siglum is in the list under "Developer Add-ins".

### Word on the web

1. Open a document in Word on the web.
2. Open Home > Add-ins > More Add-ins.
3. On the "My Add-ins" tab, click "Upload My Add-in".
4. Select `manifest.xml` and click Upload.

The browser keeps the add-in for this document. Upload the manifest again when you use another browser.

## Where to host the files

The pane is a web page, and Word loads it from the address in the manifest. You must put the files on a web server that the people who use the add-in can reach. Any static host works, because Siglum has no server program of its own. The sources stay in the document, and the page sends them nowhere. A DOI or PMID lookup goes from your browser to doi.org or to PubMed.

Build the files for the address of your host:

```sh
npm ci
npm run build -- --base https://siglum.example.org
```

Upload the contents of `dist/` to that address. The `manifest.xml` in `dist/` points to it, and this is the file you sideload. The address must use HTTPS. Word for Windows also accepts `http://localhost`, which is enough to try Siglum on one computer.

## What Siglum stores in the document

Siglum keeps everything in the `.docx` file. A person who does not have the add-in sees the citations and bibliographies as normal text and can edit the document.

| Part | Form |
| --- | --- |
| Citation | Inline content control with the tag `siglum:cite:<id>` and the title `Cite: key1; key2` |
| Bibliography | Block content control with the tag `siglum:bibliography` |
| Sources, citations and settings | One custom XML part in the namespace `urn:siglum:store:1`. Its root element `<siglum>` holds JSON. |

The JSON has this shape:

```json
{
  "sources": { "smith2020": { "id": "smith2020", "type": "article-journal", "title": "…" } },
  "citations": { "<id>": [{ "id": "smith2020", "locator": "12", "label": "page" }] },
  "settings": { "locale": "en-US" }
}
```

Each source is a CSL-JSON item. `settings.locale` is the language of the citations and bibliographies. A refresh removes the entries of citations that are no longer in the document, except those of citations that wait for review as tracked deletions and those in footnotes, headers and other places outside the body. A citation that you paste into another document keeps its keys in the title, if they fit into the 64 characters that a title can have.

The code defines these names in `src/model.js`.

## Languages

The pane is in English or German. It follows the display language of Office, and you can change it under Settings. The language of the citations and bibliographies is a separate setting that Siglum stores in the document. A new document takes it from the editing language of Office.

## Current limits

- Siglum has one citation style: NLM (Vancouver) with numbers in brackets, in order of appearance. To use another style, replace the `.csl` file and its import in `src/taskpane.js`.
- Siglum finds citations in the body text only. A citation in a footnote, an endnote, a header, a footer or a text box is not numbered or refreshed: it keeps the text it had, such as `[?]`, and its sources are not in a bibliography.
- When the document has tracked changes, Siglum reads the citations from the body's XML. If that XML cannot be read, or does not hold the same citations that Word lists, Siglum changes nothing and says so in the pane. Accept or reject the tracked changes around the citations, then refresh.
- The bundled style prints locators after the number (`[1, p. 12]`) and appends notes (`, see also`); it does not print a prefix, and `suppress-author` has no visible effect in a numbered style. The NLM style file is Siglum's variant of the upstream style with locator rendering; upstream styles drop stored locators.
- A citation that you paste into another document is known there only by the keys in its title. When the keys did not fit into the title's 64 characters, the citation stays under "Unknown". A page, a note, a prefix or a hidden author does not travel with it; neither does a locator set in the pane.
- A citation that you paste into another document is known there only by the keys in its title. When the keys did not fit into the title's 64 characters, the citation stays under "Unknown". A page or a note in the citation does not travel with it.
- A comment on a typed marker is lost when the marker becomes a citation, because Word deletes a comment when all of its text is replaced.
- Only one person can use Siglum in a document at a time. When two people cite in the same document at the same moment, one set of changes can be lost.
- The pane and the citations are available in English and German.
- The end-to-end tests need Word on Windows. See [e2e/README.md](e2e/README.md).

## Development

```sh
npm ci
npm test
npm run build -- --base https://siglum.example.org
```

`npm test` runs the unit tests in Node. They do not need Word.

To try a change on Windows, build for `http://localhost:3000` and serve `dist/`:

```sh
npm run build -- --base http://localhost:3000
npm run serve
```

Then sideload `dist/manifest.xml`. For development on Windows there is a quicker way than the shared folder: add a string value under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` that holds the full path of the manifest, and start Word again.

`npm run e2e` runs the end-to-end tests. They run from WSL on a Windows computer with Word installed. They operate the pane in real documents and then examine the saved files. [e2e/README.md](e2e/README.md) says how to run them and the measurements, and what they change on Windows; [e2e/CHECKLIST.md](e2e/CHECKLIST.md) lists what a person has to check.

The code has these parts:

| File | Contents |
| --- | --- |
| `src/model.js` | The stored format, the marker parser and the plan of what to write. This file does not use Word. |
| `src/word.js` | All calls to the Word API |
| `src/render.js` | Citations and bibliographies with citeproc-js |
| `src/library.js` | Import of BibTeX, RIS and CSL-JSON |
| `src/messages.js` | The text of the pane in each language |

## License

Siglum is free software under the GNU Affero General Public License, version 3. See [LICENSE](LICENSE).

The files in `csl/` come from the [Citation Style Language](https://citationstyles.org) project and have the license CC BY-SA 3.0.

The logo is a stand-in sketch until a hand-painted one replaces it.
