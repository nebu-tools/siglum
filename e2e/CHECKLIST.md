# What a person has to check

`npm run e2e` checks in real Word for Windows what the unit tests cannot: typed markers next to fields, tracked changes, comments, no-break spaces, tabs, carets, combining accents, table cells and other add-ins' content controls; long groups and keys longer than one search; "Use wildcards" left on in Word's Find dialog; copy, cut and move of citations between sections, a tracked move, renumbering with track changes on; save, close and reopen. `node e2e/probe.mjs` shows what happens to citations in footnotes, endnotes, headers and text boxes, and to citations pasted into another document. See [README.md](README.md).

These points need a person, because Word automation cannot do them or cannot see the result.

## The clipboard and the mouse

The suite copies and moves citations with Word's object model (`Range.FormattedText`), which is what a paste does inside one document, and leaves the system clipboard alone, because other programs on the same computer use it. Check by hand, with track changes off and on:

- Select a citation, press Ctrl+C, and press Ctrl+V in the next section. Then the same with Ctrl+X.
- Drag a sentence with a citation into another section with the mouse.
- Copy a citation into another document that has the same sources, and into one that has only some of them.

After each, move the cursor: the numbers and bibliographies follow within a few seconds.

## Office in another language

With Office's display language in English and then in German, the ribbon button, its tooltip and the add-in's description show the matching texts from `manifest.xml`, and the pane opens in that language. The rig cannot change Office's display language without changing it for everyone who uses the computer.

## Working in a large document

`node e2e/measure.mjs large` records how long Word's window does not answer ("UI stalls"). Whether Word feels usable while Siglum renumbers a large document, and whether the progress line in the pane is clear, needs someone typing in it.

## Other places Word runs

The rig drives Word for Windows only. Check the pane, citing, Refresh and a bibliography by hand in Word for Mac and in Word on the web, and with two people editing one document that is stored in OneDrive or SharePoint.
