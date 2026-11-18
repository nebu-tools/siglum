# What only real Word can settle

The unit tests run against `test/fake-word.js`, which models Word's search and content controls as far as we know them. These points depend on how Word really behaves. Check each one on this branch in Word for Windows, with `npm run e2e` and by hand, and note what you see.

## Paragraph text and literal search

Siglum reads a paragraph's `text`, finds the markers in it, and then searches the paragraph for each marker's text. It converts a marker only when the number of search hits equals the number of occurrences in `text` and the range it found reads exactly like the marker. Otherwise the pane lists the marker under "Not converted". Check that the two agree, and that nothing is replaced in the wrong place, when these are near a marker in the same paragraph:

- a field, such as a date or a cross-reference, before and after the marker
- a tracked insertion that contains the marker, and a tracked deletion just before it
- a no-break space between two markers (`@a`, no-break space, `@b`): one citation
- a no-break space before a marker
- a tab between two markers: two citations
- a caret in a supplement, `@a[x^2]`: Siglum searches for `^^`
- a key typed with a combining accent (`e` followed by U+0301)
- a comment anchored on the marker

## Places

- Markers in table cells, including two cells in one row with the same marker.
- A marker typed inside an existing citation, and inside a bibliography entry: both stay text.
- A marker inside another add-in's content control that is inside a bibliography: it stays text.
- A citation whose text was edited to hold `@a`, with `@b` typed after a space: Siglum replaces nothing and reports `@a @b` as not converted. The same with `@b` before the citation.
- A citation inside another add-in's content control, with markers before, inside and after the outer control.
- Several markers of different lengths in one paragraph, converted in one refresh: each citation stands where its marker stood.

## Long groups

- A group of 25 markers with supplements, more than 255 characters, cited as one citation. Siglum finds its first and last part and takes the range between them.
- The same long group twice in one paragraph.
- Two different long groups in one paragraph that begin or end with the same markers.
- A key of 130 letters outside the basic plane, such as rare Chinese characters, which is over 255 UTF-16 units: the search does not fail.
- A supplement with `^p`, `^t`, `^?` and `*` after "Use wildcards" was switched on in Word's Find dialog: the marker is still found as typed.

## Save and reopen

- In a new document with no store, import a source, cite it, save, close and reopen: the sources and the citation are there.

- Convert markers, save, close, reopen the document, and refresh twice. The first refresh writes nothing, and the second neither.

## Size

- A document with 4,000 citations, ten bibliographies and a few hundred typed markers. Note how long a refresh takes and whether Word stays usable. A refresh makes a fixed number of `context.sync` calls; the time should grow with the document, not with the square of it.

## Manifest

- With Office in English and in German, the ribbon button, its tooltip and the add-in's description show the matching texts from `manifest.xml`.
