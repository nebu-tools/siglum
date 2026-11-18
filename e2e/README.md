# Word rig

Everything here runs the add-in in real Word for Windows, driven from WSL on the same computer. The pages run in the task pane next to the real pane and use it through its interface, as a person does.

| File | What it is |
| --- | --- |
| `run.mjs` + `page.js` | The end-to-end suite (`npm run e2e`): pass or fail per check, in seven scenarios. |
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

Run the measurements in the background. Step groups for `--steps`: `idle`, `watcher`, `cite`, `marker`, `tracked`, `full`, `search`, `reads`; opening the pane is always measured. Every harness rebuilds and wipes its output directory, `dist-e2e/` unless `E2E_OUT` names another; give a second harness that runs at the same time its own `E2E_OUT` and `E2E_PORT`.

`measure.mjs` writes `dist-e2e/measure-<size>.json` with every row and prints a line per operation:

```
cite from search at the start (all numbers shift): 24047 ms, 15 syncs (23460 ms in Word: 19 15 2 …) writes 334c/1b/0r, heap 40 MB
```

`ms` is the time from the click or keypress until the pane reports done. `syncs` counts the pane's own `context.sync` calls in that time, and the list after the colon is how long each one waited for Word. `writes` is what the operation wrote: citations, bibliographies, removed controls. The `read:` rows time each way of reading the controls' texts on its own.

## What the rig changes on Windows

Each Word session, in `word.ps1`:

- adds one string value under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, named after the harness page's add-in id, and removes it at the end; it creates the `Developer` key if it is missing and removes it again if it is then empty;
- creates a folder `%TEMP%\siglum-e2e-<pid>-<random>` with the manifest, the document and the commands, and `rig.mjs` deletes it at the end; `probe.mjs` does the same with `%TEMP%\siglum-co-<pid>`;
- starts a Word process of its own and quits it at the end. It checks that Word started a new process and that the document opened in it; otherwise it stops and leaves the other Word alone. If its Word does not quit within 30 seconds, it stops that process by its id. Other Word windows are never touched;
- opens and saves without adding to Word's list of recent files, and removes the entry Word makes under `HKCU\Software\Microsoft\Office\16.0\Word\Reading Locations` for the document.

Word itself keeps a copy of each sideloaded manifest in its add-in cache, `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, as it does for every add-in; the rig does not remove it. The suite turns Word's "Use wildcards" find option on and off again in its own Word instance.

To check that nothing is left:

```sh
powershell.exe -NoProfile -Command 'Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\WEF\Developer; Get-ChildItem $env:TEMP -Filter "siglum*"'
```

## Baseline: main at 475ce75, 2026-11-18

Word 16.0.20326 (Microsoft 365 Apps, Current Channel, x64), Windows 11, WebView2, WSL 2 with mirrored networking. Raw rows in `baseline/measure-*.json`, the suite's and the probe's output in `baseline/suite.txt` and `baseline/probe.txt`. Another Word instance on the computer was busy during the runs, so expect some noise.

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

On main as it was, the suite failed one check: a citation moved with track changes on kept its number at the place it came from (`baseline/suite.txt`; the `main-reopen` failure there was a wrong count in the check). This branch fixes it.

Where the time goes: the sync that reads every citation with `getReviewedText("Current")` and `("Original")` holds most of every refresh (17 s of 18 s on medium), and the sync that writes holds the rest (47 s for 1,003 citations). Large documents were measured on the earlier rig at 4,000 citations: opening the pane 357 s, a refresh with nothing to change 303 s, a full rewrite 929 s.
