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

## Cleaning up

Nothing is left behind on success or failure: the registry value is removed
in a `finally`, the temp directory is deleted, Word quits. To check:

```sh
powershell.exe -NoProfile -Command 'Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
```
