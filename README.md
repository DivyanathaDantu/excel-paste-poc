# Office paste → Excel & Word (proof of concept)

Paste from Excel or Word into a web page, keep the colours and formatting, and export to `.xlsx` or `.docx`.
It uses no paid libraries: browser APIs plus [ExcelJS](https://github.com/exceljs/exceljs) (MIT) and
[docx](https://github.com/dolanmiu/docx) (MIT), both loaded from a CDN.

## Run it

```bash
python3 -m http.server 5178
```

Open http://localhost:5178. Copy a range from Excel (or a section from Word) and paste it into the box,
or click **Excel sample** / **Word sample** to load realistic clipboard HTML.

The page compares two exports:
- **Raw clipboard HTML**: what a converter that reads only inline styles gets today.
- **Inlined at paste**: the same content after the fix.

## Why the formatting gets lost

Excel and Word keep most formatting in a `<style>` block and only reference it by class
(`<td class=xl65>`). Editors throw that block away when you paste, and a converter that reads only
`style="..."` never sees the colours.

## The fix (`src/inline-styles.js`)

In the paste handler, **before the editor touches the content**:
1. Load the clipboard HTML into a hidden, sandboxed iframe so the browser applies its CSS.
2. Read each element's computed style and write it back as an inline `style="..."`.
3. Keep Excel number formats (`mso-number-format`), which browsers drop, in `data-num-format`.

From then on every colour, font, border and number format sits on the element itself, so any converter
that reads inline styles can use it.

## Export

Both exporters read formatting through `src/read-styles.js`, which looks only at inline styles
(and `<b>`/`<i>`/`<u>` tags), the same limit as your current converter.

**Excel (`src/html-to-xlsx.js`)** writes:
- fills, font colours, bold, italic and underline
- borders (thin, medium, thick, double, dashed)
- alignment and wrapped text
- merged cells, column widths and row heights
- mixed formatting inside one cell (rich text)
- numbers, currency, percentages and dates as real numbers with Excel's own number format

**Word (`src/html-to-docx.js`)** writes:
- paragraphs and headings with fonts, sizes, colours, bold, italic, underline and strikethrough
- highlighted text (Word's highlight colour, written as run shading)
- spacing and indents, including list hanging indents
- tables with cell shading, per-side borders, merged cells, column widths, row heights and cell padding
- a landscape page when a pasted table is wider than a portrait page

## Wiring it into your editor

```js
editorElement.addEventListener('paste', async e => {
  const html = e.clipboardData.getData('text/html');   // read before any await
  if (!html) return;
  e.preventDefault();
  const inlined = await inlineOfficeHtml(html);
  // insert `inlined` through the editor's own API (e.g. its insert-HTML method)
}, true);                                             // capture phase: runs before the editor's handler
```

The editor must keep inline `style` attributes on table cells and spans. If it strips them on insert,
change its allowed-styles configuration or bypass it, or the fix is undone there instead.

### DevExtreme HtmlEditor

`src/devextreme-paste.js` is a ready-made version for `dxHtmlEditor`:

```js
import { attachOfficePaste } from './devextreme-paste.js';
const detach = attachOfficePaste(htmlEditorInstance);   // call detach() when the editor is disposed
```

Then export from `htmlEditorInstance.option('value')` with `htmlToWorkbook` / `htmlToDocument`.

Measured with DevExtreme 25.1.14, pasting the Excel sample and exporting the editor's value:

| Formatting | Plain paste today | With `attachOfficePaste` | POC page (no editor) |
|---|---|---|---|
| Cell fill colours | 0 | all | all |
| Font colours | 0 | 10 / 10 | 10 / 10 |
| Bold cells | 0 | 8 / 9 | 9 / 9 |
| Merged cells | 0 | 0 (split into separate cells) | 3 / 3 |
| Cell borders | 0 | 7 / 28 (only cells with the same border on all sides) | 28 / 28 |
| Dates as real dates in .xlsx | no | no (shown as text) | yes |

The remaining losses come from the HtmlEditor's table model, which has no merged cells, keeps one border
for the whole cell and drops custom attributes such as Excel's number formats. Keeping those needs either
exporting from a copy of the inlined HTML saved at paste time, or an editor with a full document model
(e.g. the DevExpress Rich Text Editor).

## Known limits

- Not yet tested with real clipboard data from every Office version or platform. The samples follow
  Excel's and Word's clipboard format; try your own files.
- Word highlight colours on part of a paragraph have no equivalent in an Excel cell and are dropped from the .xlsx (they are kept in the .docx).
- Images, charts, formulas and conditional-formatting rules are not carried over (only their displayed values and colours).
- Excel's default column width (8.43) shows for columns about that wide.
- Word export: Word's own numbered and bulleted lists come through as their bullet character plus indent,
  not as live Word lists. Numbers are written as displayed text, since Word has no number formats.
- macOS Quick Look does not draw table cell borders in .docx files; open the file in Word to check borders.

## Tests

```bash
npm test
```

Covers the colour, border, number-format, Word-unit and merged-cell mapping (`src/mapping.js`).

## Handover

[`docs/handover.docx`](docs/handover.docx) is a short write-up for reviewers: why formatting is lost,
the libraries used, how to replicate the POC, how to wire it into DevExtreme, measured results and open questions.

## Licence

[MIT](LICENSE)
