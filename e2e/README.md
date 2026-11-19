# Word rig

Everything here runs the add-in in real Word for Windows, driven from WSL on the same computer. The pages run in the task pane next to the real pane and use it through its interface, as a person does.

| File | What it is |
| --- | --- |
| `run.mjs` + `page.js` | The end-to-end suite (`npm run e2e`): pass or fail per check, in ten scenarios. |
| `measure.mjs` + `measure.js` | Timings on generated documents: opening the pane, refresh, citing, a typed marker, the selection watcher, track changes, a full rewrite. |
| `probe.mjs` + `probe.js` | Findings, not checks, about what the README lists as limits: citations in footnotes, endnotes, headers and text boxes, citations pasted into another document, two Word instances on one file. |
| `report.mjs` | Puts `measure-*.json` files side by side as a Markdown table. |
| `generate.mjs` | Builds a document with pages, citations, bibliographies and sources, numbered by the add-in's own `plan()` and citeproc. |
| `rig.mjs` | Builds the pages, makes docx files that open a pane by themselves, runs one document in a Word instance of its own. |
| `pane.js` | What the pages share: waiting for the pane, counting `context.sync` calls, reading the document, commands to Word. |
| `word.ps1` | Starts one Word instance, sideloads the add-in, runs the pages' commands, saves, quits, and records Word's build, memory and the times its window did not answer. |
| `coauthor.ps1` | Opens one file in two Word instances, for `probe.mjs`. |
| `CHECKLIST.md` | What only a person can check, and why. |

The rig reads every stored name (tags, title prefix, namespace) from `src/model.js`.

## Scenarios of the suite

| Scenario | What it does |
| --- | --- |
| `main` | An empty document without a store: import, search, citing, typed markers, two bibliographies, an appendix and a table cell, joining a citation, renumbering, deleting, a pasted copy, track changes, DOI and PMID lookup. |
| `main-reopen` | `main`'s saved result opened again: the sources and citations are there, and two refreshes write nothing. |
| `template` | Headings with placeholder paragraphs, one of them a list item: one bibliography per section, a bibliography inserted with track changes on. |
| `template-reopen` | `template`'s result opened again, and citing goes on. |
| `markers` | Typed markers next to fields, in tracked insertions, after tracked deletions, with comments, no-break spaces, tabs, carets and combining accents; in table cells, inside citations, bibliographies and other add-ins' controls; groups longer than one search; a key of 130 letters outside the basic plane; Word's "Use wildcards" left on. |
| `sections` | Copy, cut and paste of citations between sections, a paragraph moved with track changes on (Alt+Shift+Down), renumbering with track changes on. |
| `sections-reopen` | `sections`' result opened again; after accepting the changes, a refresh writes nothing. |
| `renumber` | A generated document of 300 citations: a new first citation is numbered at once and the rest follows in batches with a progress line; a second citation and a deletion while the batches run; the pane closed (reloaded) half-way; the document saved and closed half-way. Afterwards a refresh writes nothing, and the stored entries of deleted citations are gone. |
| `renumber-reopen` | `renumber`'s result, saved while the batches ran, opened again: the first refresh writes the rest, the next one nothing. |
| `fastpath` | A generated document of 600 citations: a new first citation renumbers them in one body write after the bibliographies; a text box, a person's bookmark and a comment thread pass through. |

A check that fails because of a defect listed in the README's limits is reported as `KNOWN` with the reason and does not fail the run.

## Needs

- WSL 2 on Windows with Word for Windows (tested: Microsoft 365 Apps, 16.0.20326, Current Channel, x64, WebView2 pane), and mirrored networking (`networkingMode=mirrored` in `.wslconfig`), so that Word reaches the server in WSL at `localhost`.
- Node in WSL (tested with 26.13), `zip` and `unzip`.
- For the DOI and PMID checks in `main`, internet access.

No certificate, loopback exemption or trusted catalog is needed. The pane loads from `http://localhost:3187`, served from WSL by `scripts/serve.mjs`; Word for Windows accepts plain http for localhost. Set `E2E_PORT` for another port.

## Commands

From the repository root:

```sh
npm ci && npm test                      # unit tests
npm run e2e                             # the suite, about 3 minutes; results in dist-e2e/<scenario>.docx
npm run e2e -- markers sections         # some scenarios
node e2e/probe.mjs                      # findings, about 1 minute
node e2e/measure.mjs small              # 20 pages, 100 citations, 1 bibliography, 50 sources: 1 minute
node e2e/measure.mjs medium             # 150 pages, 1,000 citations, 3 bibliographies, 500 sources
node e2e/measure.mjs large              # 500 pages, 4,000 citations, 6 bibliographies, 2,000 sources
node e2e/measure.mjs large --steps idle,cite
node e2e/measure.mjs --pages 50 --citations 300 --bibs 2 --refs 200
node e2e/report.mjs                     # table from dist-e2e/measure-*.json
node e2e/report.mjs main=e2e/baseline/measure-medium.json now=dist-e2e/measure-medium.json
node e2e/generate.mjs --pages 20 --citations 100 --bibs 1 --refs 50 out.docx   # only the document
```

Run the measurements in the background. Step groups for `--steps`: `idle`, `watcher`, `cite`, `marker`, `tracked`, `full`, `search`, `reads`, and `floor`, which runs only when named; opening the pane is always measured. Every harness rebuilds and wipes its output directory, `dist-e2e/` unless `E2E_OUT` names another; give a second harness that runs at the same time its own `E2E_OUT` and `E2E_PORT`.

`measure.mjs` writes `dist-e2e/measure-<size>.json` with every row and prints a line per operation:

```
cite from search at the start (all numbers shift): 24047 ms, 15 syncs (23460 ms in Word: 19 15 2 …) writes 334c/1b/0r, heap 40 MB
```

`ms` is the time from the click or keypress until the pane reports done. `syncs` counts the pane's own `context.sync` calls in that time, and the list after the colon is how long each one waited for Word. `writes` is what the operation wrote: citations, bibliographies, removed controls. The `read:` rows time each way of reading the controls' texts on its own.

## What the rig changes on Windows

Each Word session, in `word.ps1`:

- sets one string value under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, named after the harness page's add-in id, and restores any pre-existing value of the same name (with its original type and content) at the end, or removes the value if it did not exist; it creates the `Developer` key if it is missing and removes it again if it is then empty;
- creates a folder `%TEMP%\siglum-e2e-<pid>-<random>` with the manifest, the document and the commands, and `rig.mjs` deletes it at the end; `probe.mjs` does the same with `%TEMP%\siglum-co-<pid>`;
- starts a Word process of its own and quits it at the end. It verifies the Word process id from a window of the application it started before opening the document, and checks that the document opened in the same process; otherwise it stops and leaves the other Word alone. If its Word does not quit within 30 seconds, it stops that process by its id. Other Word windows are never touched;
- opens and saves without adding to Word's list of recent files, and removes the entry Word makes under `HKCU\Software\Microsoft\Office\16.0\Word\Reading Locations` for the document.

Word itself keeps a copy of each sideloaded manifest in its add-in cache, `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, as it does for every add-in; the rig does not remove it. The suite turns Word's "Use wildcards" find option on and off again in its own Word instance.

To check that nothing is left:

```sh
powershell.exe -NoProfile -Command 'Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\WEF\Developer; Get-ChildItem $env:TEMP -Filter "siglum*"'
```

## Baseline, 2026-11-18

Word 16.0.20326 (Microsoft 365 Apps, Current Channel, x64), Windows 11, WebView2, WSL 2 with mirrored networking. Another Word instance on the computer was busy during the runs, so expect noise of a few tens of percent. Raw rows are in `baseline/measure-*.json`, the floor experiment in `baseline/measure-large-floor.json`, the suite's and the probe's output in `baseline/suite.txt` and `baseline/probe.txt`; `baseline/main/` holds the same for main before the changes.

For an operation that renumbers, `ms` is the time until the pane reports the operation done, `visible` until the new citation at the cursor has its number, and `settled` until the whole document is current.

### Before and after

| operation | medium, main | medium, now | large, main (earlier rig) | large, now |
| --- | --- | --- | --- | --- |
| open the pane | 18 s | 1.7 s | 357 s | 4.5 s |
| refresh, nothing to change | 18 s | 0.3 s | 303 s | 1.3 s |
| the watcher's refresh after a deleted citation | 22 s | 2.1 s | 377 s | 3.5 s |
| cite from search at the start: number visible / document current | 24 s | 0.4 s / 3.5 s | 450–520 s | 1.2 s / 16 s |
| cite from search near the end | 29 s | 0.3 s / 0.6 s | 450–520 s | 1.1 s / 2.2 s |
| typed marker at the start + refresh | 47 s | 3.7 s | 564 s | 20 s |
| cite at the start with track changes on | 49 s | 0.3 s / 4.8 s | 446 s | 1.2 s / 36 s |
| refresh with a tracked change in the document | 29 s | 0.9 s | – | 3.3 s |
| full rewrite of every citation | 71 s | 36 s | 929 s | 342 s |

### Now

| operation | small (16 p., 100 cit., 1 bibl., 50 sources) | medium (142 p., 1000 cit., 3 bibl., 500 sources) | large (450 p., 4000 cit., 6 bibl., 2000 sources) |
| --- | --- | --- | --- |
| open pane | 0.8 s, 10 syncs, 15 % in Word, writes 0c/0b | 1.7 s, 11 syncs, 35 % in Word, writes 0c/0b | 4.5 s, 12 syncs, 58 % in Word, writes 0c/0b |
| refresh, nothing to change | 0.1 s, 10 syncs, 118 % in Word, writes 0c/0b | 0.3 s, 10 syncs, 119 % in Word, writes 0c/0b | 1.3 s, 9 syncs, 124 % in Word, writes 0c/0b |
| signature (watcher's check) | 0.0 s, 1 syncs | 0.1 s, 1 syncs | 0.4 s, 1 syncs |
| cursor move, nothing changed (3.5 s window) | 3 syncs, no refresh | 2 syncs, no refresh | 2 syncs, no refresh |
| cursor move after a deletion (watcher refresh) | 1.3 s, 9 syncs, 8 % in Word, writes 0c/0b | 2.1 s, 14 syncs, 36 % in Word, writes 77c/1b | 3.5 s, 9 syncs, 54 % in Word, writes 0c/0b |
| cite from search at the start (all numbers shift) | 0.8 s, visible 0.2 s, settled 1.5 s, 21 syncs, 161 % in Word, writes 99c/1b | 1.2 s, visible 0.4 s, settled 3.5 s, 33 syncs, 284 % in Word, writes 333c/1b | 3.3 s, visible 1.2 s, settled 16 s, 87 syncs, 459 % in Word, writes 667c/1b |
| cite from search near the end | 0.3 s, visible 0.2 s, 13 syncs, 86 % in Word, writes 0c/1b | 0.6 s, visible 0.3 s, 13 syncs, 86 % in Word, writes 0c/1b | 2.2 s, visible 1.1 s, 13 syncs, 95 % in Word, writes 0c/1b |
| typed marker at the start + refresh | 0.7 s, settled 1.5 s, 25 syncs, 206 % in Word, writes 102c/1b | 1.0 s, settled 3.7 s, 37 syncs, 367 % in Word, writes 335c/1b | 3.3 s, settled 20 s, 115 syncs, 586 % in Word, writes 669c/1b |
| cite at the start with track changes on | 1.0 s, visible 0.2 s, settled 1.8 s, 22 syncs, 179 % in Word, writes 102c/1b | 1.5 s, visible 0.3 s, settled 4.8 s, 37 syncs, 303 % in Word, writes 335c/1b | 6.4 s, visible 1.2 s, settled 36 s, 187 syncs, 551 % in Word, writes 669c/1b |
| tracked changes after that | {"Added":1} | {"Added":1} | {"Added":1} |
| refresh with a tracked change, nothing to change | 0.2 s, 12 syncs, 113 % in Word, writes 0c/0b | 0.9 s, 10 syncs, 106 % in Word, writes 0c/0b | 3.3 s, 10 syncs, 106 % in Word, writes 0c/0b |
| (setup: reset every citation text) | 1.1 s | 20 s | 219 s |
| full refresh, every citation rewritten | 0.7 s, settled 1.5 s, 20 syncs, 211 % in Word, writes 103c/0b | 2.5 s, settled 36 s, 215 syncs, 1476 % in Word, writes 1003c/0b | 6.3 s, settled 342 s, 1202 syncs, 5391 % in Word, writes 4003c/0b |
| refresh again, nothing to change | 0.2 s, 10 syncs, 104 % in Word, writes 0c/0b | 1.1 s, 10 syncs, 106 % in Word, writes 0c/0b | 3.5 s, 10 syncs, 106 % in Word, writes 0c/0b |
| search keystroke | 0.0 s | 0.0 s | 0.0 s |
| read: load tag,title,id,placeholderText | 0.0 s (104 controls) | 0.1 s (1006 controls) | 0.3 s (4009 controls) |
| read: + load text | 0.0 s (104 controls) | 0.2 s (1006 controls) | 0.7 s (4009 controls) |
| read: + getReviewedText Current | 0.2 s (104 controls) | 14 s (1006 controls) | 191 s (4009 controls) |
| read: + getReviewedText Current and Original | 0.4 s (104 controls) | 31 s (1006 controls) | 442 s (4009 controls) |
| read: count tracked changes | 0.0 s (2 tracked changes) | 0.0 s (2 tracked changes) | 0.1 s (2 tracked changes) |
| stored data | 25 KB in 1 part(s), 103 stored citations | 248 KB in 1 part(s), 1003 stored citations | 1003 KB in 1 part(s), 4003 stored citations |
| Word: open document (COM) | 0.6 s | 1.0 s | 4.4 s |
| Word: save | 0.1 s | 0.5 s | 1.4 s |
| Word: peak working set | 270 MB | 344 MB | 457 MB |
| Word: UI stalls ≥ 1 s (count, longest) | 0, 0.0 s | 0, 0.0 s | 4, 1.9 s |
| pane JS heap, last | 33 MB | 32 MB | 49 MB |
| citeproc in Node, whole document | 0.3 s | 0.9 s | 4.2 s |

### Fast path, 2026-11-19

Head `4823d04` (the fast-body-v2 branch), Word 16.0.20326 on the same computer, same 450-page document. `npm test` 133/133. E2E: 111 of 113 checks pass; the 2 failures are both the new `fastpath` scenario reading a text box back as `"shape keeps\r"` (Word ends a text frame's range with its paragraph mark — a check defect, fixed after the run), plus 2 KNOWN. Large measure, wall time 800 s:

| operation, large | 2026-11-18 baseline | now |
| --- | --- | --- |
| open pane | 4.5 s, 12 syncs | 2.4 s, 10 syncs, 36 % in Word |
| refresh, nothing to change | 1.3 s, 9 syncs | 0.7 s, 10 syncs |
| cite from search at the start: number visible / document current | 1.2 s / 16 s, 87 syncs, writes 667c/1b | 0.68 s / 8.1 s, 17 syncs, writes 667c/1b |
| typed marker at the start + refresh | 3.3 s / settled 20 s, 115 syncs | settled 9.9 s, 21 syncs, writes 669c/1b |
| cite at the start with track changes on | 1.2 s / settled 36 s, 187 syncs | 0.68 s / settled 18.6 s, 103 syncs |
| refresh with a tracked change, nothing to change | 3.3 s, 10 syncs | 2.2 s, 10 syncs |
| full refresh, every citation rewritten | settled 342 s, 1202 syncs | settled 195 s, 1112 syncs |

The cite-at-start row uses the one-write path: three `fastCites` syncs of about 0.9, 3.8 and 0.4 s. The pane reports done later than before (7.9 s against 3.3 s), because the whole write happens before done instead of in batches afterwards; the new number is visible in 0.68 s either way. The full-refresh row still went through the batches: the document held a tracked change from the earlier step, so the fast path stayed off by design. UI stalls ≥ 1 s: 6, longest 5.2 s (baseline 4, 1.9 s); the README's tens-of-percent noise warning applies.

## Where Word sets the floor

`node e2e/measure.mjs large --steps floor` runs these experiments on the large document, outside the pane (`baseline/measure-large-floor.json`):

- Writing one citation's text costs Word 15–19 ms at the start of a session, through Office.js (`insertText`) and through Word's own object model (`ContentControl.Range.Text`) alike, and the same when the text does not change. Screen updating off, draft view and background pagination off made no difference in earlier runs. In one session the cost rises: after 2,000 writes it is 25–30 ms, at the same places in the document, and clearing the undo list does not bring it back. So a new first citation that renumbers 668 citations needs at least about 10 s of Word's time, and rewriting all 4,000 citations at least 60 s, more as the session goes on; the measured 16 s and 342 s are this floor plus the batches' round trips.
- `getReviewedText` costs about 70–90 ms per control, which made the old refresh 300 s. A control's `text` costs 0.7 s for 4,000, but it follows how Word shows tracked changes, and `body.getTrackedChanges()` fails on a document with a tracked move. With tracked changes Siglum reads the body's OOXML: 1.3 s for 3.3 MB. It parses that XML with saxes, a strict parser that knows namespaces, and reads only the package's main part. In Node the parse of the large document's body (3.1 MB, 4,006 controls) takes about 0.2 s; saxes alone takes 0.15 s of it. Input that is not well-formed, or that holds other controls than Word lists, stops the operation before it writes anything.
- A bibliography of 309 entries takes 0.1 s as one OOXML insert, about 15 ms per entry when entries are rewritten one by one, and 17 s the way main wrote it: a write for every run's font and for every paragraph's indent. That was the 61 s for one bibliography.
- Right after Word opens a large document it lays out its pages, and the pane's first round trips wait for that: about 2.5 s on the large document (Word's own page count takes as long). That is most of the 4.5 s for opening the pane.

What Siglum does about it: it reads `text` and, only with tracked changes, the OOXML; it writes only citations and entries that differ; a bibliography changes entry by entry when few entries differ and in one OOXML insert when many do; a new citation gets its number from the order of the tags before anything else is read; citeproc formats one citation at a time and keeps the sections it formatted; the rest of a renumbering goes in batches of about half a second of Word's time, with "Updating citations: … of …" in the pane, unless track changes are off, nothing else needs writing and enough citations need it: then one OOXML write of the body replaces the batches, after bibliographies and leftovers went through, and every rewritten control's text is checked against the plan. A new action runs before the remaining batches. The batches of the old plan are then dropped, and Siglum plans again after the action, unless the action was itself a refresh. A failed batch is handled the same way. The pane says the document is current (`siglum:settled`) only after a plan that left nothing to write and whose order of controls still holds.

## Baseline of main at 475ce75, before the changes

Word 16.0.20326 (Microsoft 365 Apps, Current Channel, x64), Windows 11, WebView2, WSL 2 with mirrored networking. Raw rows in `baseline/main/`. Another Word instance on the computer was busy during the runs, so expect some noise.

| operation | small (16 p., 100 cit., 1 bibl., 50 sources) | medium (142 p., 1000 cit., 3 bibl., 500 sources) |
| --- | --- | --- |
| open pane | 0.7 s, 11 syncs, 28 % in Word, writes 0c/0b | 18 s, 13 syncs, 96 % in Word, writes 0c/0b |
| refresh, nothing to change | 0.3 s, 11 syncs, 87 % in Word, writes 0c/0b | 18 s, 10 syncs, 97 % in Word, writes 0c/0b |
| signature (watcher's check) | 0.0 s, 1 syncs | 0.1 s, 1 syncs |
| cursor move, nothing changed (3.5 s window) | 3 syncs, no refresh | 2 syncs, no refresh |
| cursor move after a deletion (watcher refresh) | 1.9 s, 10 syncs, 14 % in Word, writes 0c/0b | 22 s, 12 syncs, 90 % in Word, writes 77c/1b |
| cite from search at the start (all numbers shift) | 1.2 s, 15 syncs, 96 % in Word, writes 100c/1b | 24 s, 15 syncs, 98 % in Word, writes 334c/1b |
| cite from search near the end | 1.6 s, 15 syncs, 97 % in Word, writes 1c/1b | 29 s, 15 syncs, 98 % in Word, writes 1c/1b |
| typed marker at the start + refresh | 2.3 s, 17 syncs, 98 % in Word, writes 102c/1b | 47 s, 17 syncs, 98 % in Word, writes 335c/1b |
| cite at the start with track changes on | 2.2 s, 17 syncs, 97 % in Word, writes 103c/1b | 49 s, 17 syncs, 99 % in Word, writes 336c/1b |
| tracked changes after that | {"Added":1} | {"Added":1} |
| refresh with a tracked change, nothing to change | 0.6 s, 12 syncs, 92 % in Word, writes 0c/0b | 29 s, 10 syncs, 97 % in Word, writes 0c/0b |
| (setup: reset every citation text) | 1.3 s | 44 s |
| full refresh, every citation rewritten | 2.0 s, 12 syncs, 96 % in Word, writes 103c/0b | 71 s, 12 syncs, 99 % in Word, writes 1003c/0b |
| refresh again, nothing to change | 0.6 s, 10 syncs, 89 % in Word, writes 0c/0b | 29 s, 10 syncs, 97 % in Word, writes 0c/0b |
| search keystroke | 0.0 s | 0.0 s |
| read: load tag,title,id,placeholderText | 0.0 s (104 controls) | 0.1 s (1006 controls) |
| read: + load text | 0.0 s (104 controls) | 0.2 s (1006 controls) |
| read: + getReviewedText Current | 0.2 s (104 controls) | 12 s (1006 controls) |
| read: + getReviewedText Current and Original | 0.4 s (104 controls) | 29 s (1006 controls) |
| read: count tracked changes | 0.0 s (2 tracked changes) | 0.0 s (2 tracked changes) |
| stored data | 25 KB in 1 part(s), 104 stored citations | 248 KB in 1 part(s), 1004 stored citations |
| Word: open document (COM) | 0.4 s | 0.9 s |
| Word: save | 0.2 s | 0.5 s |
| Word: peak working set | 274 MB | 349 MB |
| Word: UI stalls ≥ 1 s (count, longest) | 0, 0.0 s | 0, 0.0 s |
| pane JS heap, last | 32 MB | 27 MB |
| citeproc in Node, whole document | 0.1 s | 1.8 s |

On main as it was, the suite failed one check: a citation moved with track changes on kept its number at the place it came from (`baseline/main/suite.txt`; the `main-reopen` failure there was a wrong count in the check), fixed since. A control's `text` follows how Word shows tracked changes ("All Markup" keeps deleted text in it, "Simple Markup" does not), and `getReviewedText` gives a moved text both ways, so with tracked changes in the document Siglum reads the texts from the body's OOXML. The suite checks a tracked deletion and a tracked move under both views.

Where the time goes: the sync that reads every citation with `getReviewedText("Current")` and `("Original")` holds most of every refresh (17 s of 18 s on medium), and the sync that writes holds the rest (47 s for 1,003 citations). Large documents were measured on the earlier rig at 4,000 citations: opening the pane 357 s, a refresh with nothing to change 303 s, a full rewrite 929 s.
