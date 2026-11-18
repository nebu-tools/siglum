# Word rig

Everything in this directory runs the add-in in real Word for Windows, driven
from WSL on the same machine. Nothing here changes the add-in; harness pages
run next to the real pane and use it through its UI.

| File | What it is |
| --- | --- |
| `run.mjs` + `page.js` | The end-to-end suite (`npm run e2e`): pass/fail checks in an empty document, the G-BA Modul 3 template, after `gba.R apply`, and reopened. |
| `measure.mjs` + `measure.js` | Timings on generated documents: open, refresh, cite, typed marker, selection watcher, track changes, full rewrite. |
| `probe.mjs` + `probe.js` | Findings, not checks: footnotes, endnotes, headers, text boxes, copy/cut/move, tracked move, save and reopen, paste into another document, two instances on one file. |
| `report.mjs` | Puts `measure-*.json` side by side as a Markdown table. |
| `coauthor.ps1` | Opens one file in two Word instances (used by `probe.mjs`), with a time limit and cleanup by process id. |
| `generate.mjs` | Builds a document with N pages, C citations, B reference lists and R works, rendered with the add-in's own `plan()` and citeproc. |
| `names.mjs` | The only place that knows the add-in's stored names (tags, namespace, store keys, marker syntax). It reads them from `src/model.js`. |
| `rig.mjs` | Builds the panes, makes docx files that open a pane by themselves, runs one document per Word instance. |
| `pane.js` | Browser helpers for harness pages: waiting for the pane, counting `context.sync` calls. |
| `word.ps1` | Starts one Word instance, sideloads, waits, saves, quits; records Word's build, memory and UI stalls. |

## Needs

- WSL 2 on Windows with Word for Windows (tested: Microsoft 365 Apps / Office
  2021 ProPlus Retail, 16.0.20326.20318, Current Channel, x64, WebView2 pane).
- Node in WSL (tested with 26.13; CI uses 24), `zip` and `unzip`.
- For `npm run e2e` only: a [tabellion](https://github.com/nebu-tools/tabellion)
  checkout next to this one (or `TABELLION_DIR`), `rv` on the PATH, and the
  templates: `cd ../tabellion && rv run gba.R fetch`.

No certificate, no loopback exemption and no trusted catalog are needed. The
pane loads from `http://localhost:3187`, served from WSL by
`scripts/serve.mjs`; Word for Windows accepts plain http for localhost, and
WSL forwards Windows' localhost to it (checked with `networkingMode=mirrored`
only; the default NAT mode is untested). Set `E2E_PORT` for another port.

## How sideloading works

`word.ps1` adds one string value under
`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, named after the add-in id
and pointing at a manifest in `%TEMP%\siglum-e2e-<pid>-<random>\`. It starts a
new Word instance through COM, opens the document, waits for the page to say
done, saves to `out.docx`, quits that instance and removes the value. The
document itself carries a `webextensions` part with
`Office.AutoShowTaskpaneWithDocument`, so the pane opens without a click.
Each harness page has its own add-in id (`rig.mjs` `idFor`), so Word's add-in
cache never mixes them up. Other Word windows are never touched.

## Commands

On any branch, from the repository root:

```sh
npm ci && npm test && npm run build      # unit tests and the production build
npm run e2e                              # suite; results in dist-e2e/{a,b,r,c}.docx
node e2e/measure.mjs small               # 20 pages, 100 citations, 1 list, 50 works
node e2e/measure.mjs medium              # 150 pages, 1,000 citations, 3 lists, 500 works
node e2e/measure.mjs large               # 500 pages, 4,000 citations, 6 lists, 2,000 works
node e2e/measure.mjs --pages 50 --citations 300 --bibs 2 --refs 200
node e2e/report.mjs                      # table from dist-e2e/measure-*.json
node e2e/probe.mjs                       # behaviour findings
node e2e/generate.mjs --pages 20 --citations 100 --bibs 1 --refs 50 out.docx   # just the document
```

Run measurements in the background: medium takes about 10 minutes, large about
two hours on main. Detach the large run (`setsid nohup node e2e/measure.mjs large
> large.log 2>&1 &`) if your shell kills background jobs after two hours. To
run a second harness while one is running, give it its own directory and port:
`E2E_OUT=dist-e2e-probe E2E_PORT=3188 node e2e/probe.mjs`. Every harness wipes
its output directory when it builds. `measure.mjs` writes `dist-e2e/measure-<size>.json` with every
row, and prints a line per operation:

```
cite from search at the start (all numbers shift): 31006 ms, 15 syncs (30316 ms in Word: 17 17 2 …) writes 334c/1b/0r, heap 35 MB
```

`ms` is wall time from the click or keypress until the pane reports done.
`syncs` counts the pane's own `context.sync` calls in that time, and the list
after the colon is how long each one waited for Word, in order. `writes` is
what the refresh rewrote: citations, lists, removed controls. The last lines
(`read: …`) time each way of reading the controls' text on its own.

## On a branch that renames things

If the add-in renames its exports in `src/model.js`, fix `e2e/names.mjs` and
nothing else: it tries several names per export and says which one is
missing. Store keys come from `emptyStore()`, the typed marker from whatever
`parseMarker()` accepts. `page.js`, the pass/fail suite, still has German UI
texts and `[@key]` markers in its checks; those are the assertions that have
to follow an interface translation. The harness pages listen for
`zitate:ready`/`zitate:done` and the likely renames (`siglum:*`,
`citations:*`); add the new name to `pane.js` if it differs.

## Baseline: main at 3c29799, 2026-11-18

Word 16.0.20326.20318 (Office 2021 ProPlus Retail, Current Channel, x64),
Windows 11 Pro Education build 26300, WebView2 157.0.4341.60, WSL 2 mirrored networking. Raw rows in `baseline/measure-*.json`,
probe output in `baseline/probe.txt`. Compare a branch with
`node e2e/report.mjs e2e/baseline/measure-medium.json dist-e2e/measure-medium.json`.
The large column comes from two sessions (`large`, stopped by the old
10-minute limit after the marker step, then `large --steps tracked,full,search,reads`),
and both ran while another Word instance on the machine was busy, so expect
some noise. "Pages" is Word's own count.

| operation | small (21 p., 100 cit., 1 lists, 50 works) | medium (181 p., 1000 cit., 3 lists, 500 works) | large (578 p., 4000 cit., 6 lists, 2000 works) |
| --- | --- | --- | --- |
| open pane | 1.2 s, 11 syncs, 26 % in Word, writes 0c/0b | 22 s, 13 syncs, 94 % in Word, writes 0c/0b | 357 s, 13 syncs, 100 % in Word, writes 0c/0b |
| refresh, nothing to change | 0.4 s, 11 syncs, 136 % in Word, writes 0c/0b | 36 s, 10 syncs, 97 % in Word, writes 0c/0b | 303 s, 10 syncs, 99 % in Word, writes 0c/0b |
| signature (watcher's check) | 0.0 s, 1 syncs | 0.1 s, 1 syncs | 0.3 s, 1 syncs |
| cursor move, nothing changed (3.5 s window) | 3 syncs, no refresh | 2 syncs, no refresh | 2 syncs, no refresh |
| cursor move after a deletion (watcher refresh) | 2.0 s, 10 syncs, 16 % in Word, writes 0c/0b | 32 s, 12 syncs, 94 % in Word, writes 77c/1b | 377 s, 10 syncs, 99 % in Word, writes 0c/0b |
| cite from search at the start (all numbers shift) | 1.3 s, 15 syncs, 96 % in Word, writes 100c/1b | 31 s, 15 syncs, 98 % in Word, writes 334c/1b | 521 s, 15 syncs, 99 % in Word, writes 668c/1b |
| cite from search near the end | 2.3 s, 15 syncs, 98 % in Word, writes 1c/1b | 55 s, 15 syncs, 99 % in Word, writes 1c/1b | 454 s, 15 syncs, 99 % in Word, writes 1c/1b |
| typed marker at the start + refresh | 3.0 s, 14 syncs, 99 % in Word, writes 102c/1b | 72 s, 14 syncs, 99 % in Word, writes 335c/1b | 564 s, 14 syncs, 99 % in Word, writes 669c/1b |
| cite at the start with track changes on | 1.8 s, 17 syncs, 97 % in Word, writes 103c/1b | 81 s, 17 syncs, 99 % in Word, writes 336c/1b | 446 s, 17 syncs, 99 % in Word, writes 668c/1b |
| tracked changes after that | {"Added":1} | {"Added":1} | {"Added":1} |
| (setup: reset every citation text) | 0.9 s | 41 s | 251 s |
| full refresh, every citation rewritten | 1.4 s, 12 syncs, 97 % in Word, writes 103c/0b | 90 s, 12 syncs, 99 % in Word, writes 1003c/0b | 929 s, 12 syncs, 100 % in Word, writes 4001c/0b |
| refresh again, nothing to change | 0.5 s, 10 syncs, 92 % in Word, writes 0c/0b | 43 s, 10 syncs, 98 % in Word, writes 0c/0b | 409 s, 10 syncs, 99 % in Word, writes 0c/0b |
| search keystroke | 0.0 s | 0.0 s | 0.0 s |
| stored library | 27 KB in 1 part(s) | 264 KB in 1 part(s) | 1066 KB in 1 part(s) |
| read: load tag,title,id,placeholderText |  |  | 0.4 s (4007 controls) |
| read: + load text |  |  | 0.8 s (4007 controls) |
| read: + getReviewedText Current |  |  | 210 s (4007 controls) |
| read: + getReviewedText Current and Original |  |  | 350 s (4007 controls) |
| read: count tracked changes |  |  | 0.1 s (2 tracked changes) |
| Word: open document (COM) | 0.4 s | 0.7 s | 3.5 s |
| Word: save | 0.1 s | 0.8 s | 1.1 s |
| Word: peak working set | 273 MB | 343 MB | 512 MB |
| Word: UI stalls ≥ 1 s (count, longest) |  |  | 4, 2.3 s |
| pane JS heap, last | 28 MB | 34 MB | 31 MB |
| citeproc in Node, whole document | 0.1 s | 0.8 s | 9.7 s |

Where the time goes: in every refresh one `context.sync` holds 95–99 % of the
time. It is the snapshot in `src/word.js`, which asks every control for
`getReviewedText("Current")` and `("Original")`. The `read:` rows show the
same read with `load("text")` takes 0.8 s for 4,007 controls; the two
reviewed texts take 350 s. The second large sync, when there is one, writes
the citations and lists (480 s for 4,001 citations, 61 s for one list of
about 200 entries). citeproc in the pane is a few seconds at most. Round
trips are a fixed 10–17 per operation, so their number is not the problem.

## Cleaning up

Nothing is left behind on success or failure: the registry value is removed
in a `finally`, the temp directory is deleted, Word quits. To check:

```sh
powershell.exe -NoProfile -Command 'Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
```
