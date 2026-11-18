# Experiment: can Word change thousands of citations from one write?

Branch `experiment-binding`, from `scale` at b6266d1. Word 16.0.20326 (Microsoft 365 Apps, Current Channel, x64), Windows 11, WebView2 pane, WSL 2. The machine ran other long jobs at the same time, so expect noise of some tens of percent. Nothing in `src/` changed.

## What is here

| File | What it is |
| --- | --- |
| `generate-bound.mjs` | Builds the test documents from `e2e/generate.mjs`, as OOXML: `plain` (plain-text controls bound to a custom XML part), `rich` (rich-text controls bound with `w15:dataBinding`), `fields` (REF fields to bookmarked, numbered bibliography entries), `unbound` (the generated document with the experiment's tags), `plain-unbound` (plain-text controls without binding), `features` (a small document with a bound control in every place a citation can be) and `mini`. |
| `shared.js` | Names and helpers that the generator and the page share. |
| `bind.mjs` | Runs one session: `node e2e/experiments/bind.mjs <session>`. Writes `readings/<session>.json`, and `readings/<session>.partial.jsonl` while it runs. |
| `bind.js` | The page in the pane: the steps `open`, `timing`, `undo`, `tracked`, `reopen`, `fields`, `direct`, `paths`, `bodyooxml`, `features`, `limits`, `mini`. |
| `word-exp.mjs`, `word-exp.ps1` | Copies of the rig's `inWord()` and `word.ps1`. They find their own Word process from a window of the application they created (a blank document) before opening the test document. They put back a registry value of the same name if one was there before. |
| `readings/` | The raw readings, one file per session. `rich.json` is the session whose Word crashed while opening; `fields-large-run1.partial.jsonl` and `rich-medium.partial.jsonl` are sessions that stopped before the end (a restart of the host, and a crash of Word). `probe*.json` are the first passes that found the fast and the slow write paths. |

The tags are `exp:cite:<id>` and `exp:bib`, so the real pane next to the page leaves the controls alone. The documents ask Word to remove personal information on save (`w:removePersonalInformation`): the saved files have no author, no company and no last-modified-by. In the readings, a revision author shows only as "Author B" or "Author", or as "(another name)" when it is the name of the account signed in to Word.

## Results

A "run" is one measurement in a session. "Time" is from the Office.js call until every control in the document shows the new text, checked by reading all controls through Office.js right after the call, then through Word's object model (COM), then in the saved file. On a large document, `setXml` returns before Word is done: Word keeps updating the controls for minutes afterwards and blocks every other call meanwhile.

### Times, 4,000 citations over 450 pages

| candidate, write path | all 4,000 changed | 10 % (400) | one | runs |
| --- | --- | --- | --- | --- |
| A, plain-text binding, `CustomXmlPart.selectSingleNode(xpath).text` per changed element, one sync (WordApiDesktop 1.3) | 44.9, 44.3, 42.9 s | 4.6, 5.4, 4.5 s | 0.02, 0.15, 0.01 s | 3 (`plain`); 1 more each in `probe-pos` |
| A, the same with positional xpaths | 40.1 s | 5.2 s | 0.02 s | 1 (`probe-pos`) |
| A, `selectNodes("/*/*")` and set `.text` (WordApiDesktop 1.3) | 54 s | 15.5 s | 11 s | 1 (`probe`) |
| A, common API `CustomXMLNode.setTextAsync`, one call after the other (CustomXmlParts 1.1) | 85 s | 9.0, 11.2, 9.0 s | 0.21, 0.47, 0.15 s | 3, all once (`plain`, `probe`) |
| A, `CustomXmlPart.setXml` (whole part, WordApi 1.4) | 908 s | 113 s | 21 s | 1 each (`probe`, first `plain` attempt) |
| A, `CustomXmlPart.updateElement` (WordApi 1.4) | not run: ~3 s per element at 1,000 citations, 9–19 s at 4,000 | not run | 3–9 s | 1 (`probe-medium`, `probe-pos-medium`) |
| A, plain-text binding, `selectSingleNode().text`, **20,000 citations** | 1,197 s | 119 s (2,000) | 0.74 s | 1 (`plain-20k`) |
| A, `setXml`, 20,000 citations | not run | not run | 630 s | 1 (`plain-20k`) |
| A, rich-text binding | Word crashed on opening the 4,000-citation document, twice. At 1,000: 87 ms per control, and writing all of them crashed Word. | | | 2 (`rich`), 1 (`rich-medium`) |
| B, REF fields: Office.js `field.updateResult()` on every REF field, one sync (WordApi 1.5) | 957 s (4,998 fields) | 1,253 s (every field updated, 239 citations changed) | not finished | 1 (`fields-large-run1.partial.jsonl`) |
| B, at 1,000 citations: Word's own `Fields.Update()` through COM (no add-in can call it) | 47 s (1,247 fields) | – | – | 1 (`fields-medium`) |
| C, the current way: `insertText` into each rich-text control, one sync | 106, 239, 341 s | 6.1, 18.4, 29.4 s | 0.03, 0.03, 0.06 s | 3, one session (`unbound`) |
| C, `insertText` into unbound plain-text controls | 106, 230, 357 s | 6.6, 16.5, 30.6 s | 0.03, 0.02, 0.07 s | 3, one session (`plain-unbound`) |
| C, one `body.getOoxml()` + `body.insertOoxml(…, "Replace")` with every citation text changed, track changes off | 5.5, 6.2, 6.9 s | (the same: the whole body is written) | (the same) | 3 (`body-ooxml`) |
| C, one `insertOoxml` over paragraphs 20–400 (609 citations), track changes on | 60 s, and the old content stays as a tracked deletion | | | 1 (`unbound`) |

In one session the direct write gets slower, from 15–27 ms per control at the start to 85–90 ms after about 13,000 writes, both for rich-text and plain-text controls. The bound write stayed at 11 ms per control over three runs of 4,400 writes each, and it stayed there under track changes too. The body OOXML write grew from 5.5 to 6.9 s over three runs.

Per control, late in the `unbound` session: `insertText` 98–102 ms, `getRange("Content").insertText` the same, `insertOoxml` 222 ms, COM `Range.Text` 92 ms.

### Size and open time (point 10)

| document, 4,000 citations | file | Word opens it | save | first pane read of all controls | Word's peak memory |
| --- | --- | --- | --- | --- | --- |
| unbound (rich-text controls, as now) | 565 KB | 2.5 s | 0.8–1.1 s | 0.47–0.57 s | 390–505 MB |
| plain-text controls, part present, no binding | 585 KB | 2.6 s | 0.8 s | – | 489 MB |
| plain-text controls bound to the part, id xpaths | 618 KB (730 KB after Word saved it) | 23.8–24.8 s | 16.9 s | 0.59–0.61 s | 530–938 MB |
| the same, positional xpaths | 615 KB | 20.8 s | 12.5 s | – | 960 MB |
| 20,000 plain-text controls bound | 1.26 MB | 679 s | 404 s | 2.1 s | 649 MB |
| rich-text controls bound | 1,800 KB part | Word crashed while opening, twice | – | – | – |

The binding, not the plain-text control or the part, costs the 20 s on opening and the 12–16 s on saving. Word's UI froze for 10–110 s at a time during `setXml` and `updateElement` (`stalls` in the readings).

## Candidate A: content controls bound to a custom XML part

1. **When is it done?** Tested. With `selectSingleNode(...).text` and the common API, every control reads right through Office.js and COM, including the last citations far from the view, as soon as the call returns. With `setXml` the call returns after 13–31 s and Word then updates the changed controls for up to 15 minutes, at about 60–240 ms each, blocking every call. In the saved file, document.xml holds the new text in every control (0 of 4,000 differ from the part).
2. **Save, close, reopen.** Tested and seen. The text holds. document.xml carries the current text, so other tools read it. On opening, Word takes the part's value: three citations changed only in the part showed the part's text, and three changed only in document.xml also showed the part's text (`plain-reopen`, `plain-reopen-doc`). One of two attempts to open the edited file failed with "The operation is canceled"; the second opened.
3. **Track changes.** Tested and seen. A bound update with track changes on is recorded as a revision, a deletion and an insertion per control (800 revisions for 400 controls), under the name of the account signed in to Word. It was neither the Office user name from the Options nor empty. `control.text` then reads old and new together ("[15.0][t.0]"); `getReviewedText` gives both. Rejecting puts the old text back in the control and in the part; accepting keeps the new text; the binding stays in both cases. A typed tracked change in a bound control reaches the part at once; rejecting it puts the element back. `body.getTrackedChanges()` works here (800 items), except once a tracked move is in the document (GeneralException, as known).
4. **Typing and locks.** Tested and seen. A person typing in a plain-text bound control writes the text into the element at once, and other controls bound to that element follow. With `w:lock` `sdtContentLocked` or `contentLocked`, typing is refused ("The Insert statement is currently disabled"), and so is Office.js `insertText` (GeneralException), but a part write still updates the control. With `sdtLocked` only, the person can type (and the part follows) but not delete the control.
5. **Formatting.** Tested and seen. Superscript set as run properties on the control (`w:sdtPr/w:rPr`) stays after a bound update (`Font.Superscript` = true). Plain-text binding carries text only. Rich-text binding needs `w15:dataBinding`; `w:dataBinding` on a rich-text control shows the escaped package as text. With `w15:dataBinding` the element holds an escaped flat-OPC package, and an italic "et al." in the middle shows italic after a part write. Office.js can write it (`node.text = package`, `setXml`). It costs about 1 KB per citation in the part, 87 ms per control update at 1,000 citations (against 11 ms for plain text), and a person's edit makes Word store a whole package with the style definitions in the element (6.5 KB). Word crashed opening 4,000 of them (twice) and while writing 1,000 (once). Out.
6. **Copies.** Tested through Word's `Range.FormattedText`, which copies inside Word without the clipboard. The clipboard of the shared machine held someone's content, so Ctrl+C/Ctrl+V, cut and paste and drag were not tested. A copy in the same document keeps the binding and follows later part writes. A copy into another document keeps the xpath but is not mapped there, and the part does not come along: the text is frozen. A cut-and-paste (copy, then delete the original) keeps the binding. A paragraph moved with track changes on (`Relocate`) gives a tracked move; both the old and the new copy stay bound and follow the part.
7. **Places.** Tested and seen. Bound controls in a footnote, a header, a text box, a table cell and a comment all show a part write the same way. Office.js does not list plain-text controls in `body.contentControls`, `document.contentControls`, `paragraph.contentControls` or `getByTag`; only `body.getContentControls()` (WordApi 1.5) returns them. Siglum reads controls through `body.contentControls`, so a move to plain-text controls must change its reads.
8. **Two controls bound to one element.** Tested and seen. Both show every part write and every typed change in either.
9. **Undo.** Tested. A `setXml` write does not go on the undo list (`Undo` returns false). A write of 400 elements through `selectSingleNode().text` made two undo steps: the first undid 3 of the 400, the second the remaining 397; control text and part went back together. The body OOXML write is one undo step.
10. **Size and open time.** See the table above.
11. **Limits.** Tested on Word for Windows: `customXmlParts.add`, `getXml` and `setXml` of parts of 1, 5, 20, 50 and 100 MB all worked, taking 0.04 to 3.4 s (`limits`). No limit was met. I found no documented payload limit for Word; Microsoft documents a 5 MB request limit only for Excel on the web. The limit that matters is Word's time: with 20,000 bound controls Word took 679 s to open the document (11 minutes, against 24 s for 4,000), a bound write cost 60 ms per control (11 ms at 4,000), and one `setXml` changing one value took 630 s. Word's cost per bound write grows with the number of bindings in the document.
12. **Co-authoring.** Not run. Co-authoring needs a file on OneDrive or SharePoint, and `e2e/coauthor.ps1` may not be used as it is.
13. **Word on the web and on Mac.** Read, not tested. `XmlMapping`, `CustomXmlNode` and `selectSingleNode` are in WordApiDesktop 1.3. Microsoft lists that set for Windows (version 2507) and Mac (16.99.2) and not for the web ([Word API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/word/word-api-requirement-sets)). `setXml` and `updateElement` are in WordApi 1.4, which the web has, but they are the slow paths. On the web, Microsoft says content controls are kept but "cannot be edited or updated" in editing view ([differences between the browser and Word](https://support.microsoft.com/en-us/word/differences-between-using-a-document-in-the-browser-and-in-word)). I found nothing that says whether Word on the web or on Mac redraws a bound control after a part write.

## Candidate B: REF fields to numbered bibliography entries

1. **One call for all fields?** No. Office.js has `field.updateResult()` (WordApi 1.5), per field. In one sync over 4,998 REF fields it took 957 s (191 ms per field); 100 fields at 1,000 citations took 14 ms each. Word's own `Fields.Update()` is COM only: 47 s for 1,247 fields. Word does not update REF fields on open; "update fields before printing" was not tried, because it is an application setting. Moving bibliography entries is cheap: one OOXML insert per bibliography, with the bookmarks in it, took 1.4–2.8 s for 1,894 entries.
2. **Grouped citations.** Fields cannot make "[1–3]". The document shows "[5,3]" and "[9,10]": one field per source, in the citation's order, not sorted and not collapsed. After reversing a bibliography the first entry read "[0]" until the next update. This candidate serves only citations of one source each.
3. **Track changes, copies, and the time between.** With track changes on, an update of all fields recorded 2,478 revisions (two per changed number). A citation copied into another document reads "[1,221]" until its fields update, then "[Error! Reference source not found.,Error! Reference source not found.]". Between a reorder and the update the reader sees the old numbers.

## Candidate C: other paths

- **One OOXML write of the whole body**, with every citation text changed in the OOXML: 5.5–6.9 s for all 4,000, track changes off. Afterwards Word's object model showed the same control tags and ids, the same comment (text, scope, date), the same tracked insertion and deletion (type, text, date), the same footnote, the same PAGE field and the same paragraph count. One defect: the bookmark that covered paragraph 30 grew by one character (506 → 507, then 509 after three writes). With track changes on, an `insertOoxml` over a range is recorded as a deletion of the whole range and an insertion of the new one: 60 s for 609 citations, and the document then held both copies (4,615 controls, two comments). With track changes on this path is out. What I did not check: modern comment threads with replies, images and other embedded objects (they make the OOXML large), content controls of other add-ins, the cursor and scroll position after the write, and two people co-authoring.
- **Faster ways to write a control's text:** none. `insertText`, `getRange("Content").insertText`, writing only the first character, and COM `Range.Text` cost the same per control; `insertOoxml` per control costs twice as much.
- **Fields from a document property or a document variable** were not built. They need a field update per field, like B.

## Conclusion

- No path makes Word change thousands of citations for the cost of one write while track changes are on. Word updates each bound control by itself, at about 11 ms per control with 4,000 bindings and 60 ms with 20,000. Every changed control becomes two revisions.
- With track changes off, one OOXML write of the body changes all 4,000 citations in 5.5–6.9 s and kept everything checked in it except one bookmark's range.
- The binding avoids the slowdown of direct writes within a session (43–45 s in every run, against 106 s, then 239 s, then 341 s). It costs about 20 s more each time Word opens a 4,000-citation document and 12–16 s more each time Word saves it; at 20,000 citations, 11 minutes to open. It also takes plain-text controls out of `body.contentControls`.

## What it changes on Windows

Each session (`word-exp.ps1`) adds one string value under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` and removes it at the end, together with the key if it made it. It also creates `%TEMP%\siglum-e2e-<pid>-<random>`, which `word-exp.mjs` deletes, starts one Word process and quits it. When Word crashes or is stopped, Word itself adds entries under `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency` (`DisabledItems`, `DocumentRecovery`) that name the session's `in.docx`, and it can leave an AutoRecover file in `%APPDATA%\Microsoft\Word`. Remove those by hand, matching the exact session folder.
