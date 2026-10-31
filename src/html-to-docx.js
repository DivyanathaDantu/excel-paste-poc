// HTML -> Word document with the `docx` library. Formatting is read from inline styles only (see read-styles.js).

import { cssLengthToPx, cssLengthToTwips, docxBorderFromCss, horizontalAlign, layoutTable } from './mapping.js';
import { cellFormat, collectRuns, fontFromStyle, newStats, walkContent } from './read-styles.js';

const USABLE_PORTRAIT_TWIPS = 9360;   // 8.5in page minus 1in margins
const DEFAULT_COLUMN_PX = 96;

// docx keeps table options private, so remember each table's width for the page orientation check
const tableWidths = new WeakMap();

/** Removes undefined keys, so docx only writes properties the source actually set. */
const defined = obj => Object.fromEntries(Object.entries(obj).filter(([, value]) => value !== undefined));
const definedOrNone = obj => {
  const result = defined(obj);
  return Object.keys(result).length ? result : undefined;
};

const shading = (docx, fill) => ({ type: docx.ShadingType.CLEAR, color: 'auto', fill });

function textRuns(runs, baseFont, docx, stats) {
  const out = [];
  for (const run of runs) {
    const font = { ...baseFont, ...run.font };
    const props = defined({
      font: font.name,
      size: font.size ? Math.round(font.size * 2) : undefined,   // docx sizes are half-points
      color: font.color,
      bold: font.bold,
      italics: font.italic,
      underline: font.underline ? {} : undefined,
      strike: font.strike,
      shading: run.background ? shading(docx, run.background) : undefined,
    });
    if (run.background) stats.highlights++;
    if (font.color && font.color !== '000000') stats.fontColors++;
    if (font.bold) stats.bold++;
    run.text.split('\n').forEach((line, i) => {
      out.push(new docx.TextRun({ ...props, text: line, ...(i > 0 ? { break: 1 } : {}) }));
    });
  }
  return out;
}

function paragraph(el, docx, stats) {
  const style = el.style;
  const runs = collectRuns(el);
  const textIndent = cssLengthToTwips(style.textIndent);
  const level = /^H([1-6])$/.exec(el.tagName)?.[1];
  const alignment = horizontalAlign(style.textAlign || el.getAttribute('align'));

  if (runs.length > 1) stats.richText++;
  stats.paragraphs++;
  return new docx.Paragraph(defined({
    children: textRuns(runs, fontFromStyle(style), docx, stats),
    heading: level ? docx.HeadingLevel[`HEADING_${level}`] : undefined,
    bullet: el.tagName === 'LI' ? { level: 0 } : undefined,
    alignment: alignment === 'justify' ? docx.AlignmentType.JUSTIFIED : alignment,
    spacing: definedOrNone({ before: cssLengthToTwips(style.marginTop), after: cssLengthToTwips(style.marginBottom) }),
    indent: definedOrNone({
      left: cssLengthToTwips(style.marginLeft) || undefined,
      firstLine: textIndent > 0 ? textIndent : undefined,
      hanging: textIndent < 0 ? -textIndent : undefined,
    }),
  }));
}

/** Content of a container (body or table cell) as docx paragraphs and tables, in order. */
function blocks(container, docx, stats) {
  const out = [];
  walkContent(container, {
    onTable: table => out.push(tableToDocx(table, docx, stats), new docx.Paragraph({})),
    onBlock: el => out.push(paragraph(el, docx, stats)),
  });
  return out;
}

function cellBorders(format, stats) {
  const borders = {};
  for (const [side, css] of Object.entries(format.borders)) {
    const border = docxBorderFromCss(css.width, css.style, css.color);
    if (border) borders[side] = border;
  }
  if (Object.keys(borders).length) stats.borders++;
  return borders;
}

function tableToDocx(table, docx, stats) {
  const rows = [...table.rows];
  const { cells, columnCount } = layoutTable(
    rows.map(tr => [...tr.cells].map(td => ({ colspan: td.colSpan, rowspan: td.rowSpan }))),
  );

  // Column widths from single-column cells (Excel declares them on the first row only)
  const columnPx = Array(columnCount).fill(null);
  for (const placed of cells) {
    if (placed.colspan !== 1) continue;
    const td = rows[placed.rowIndex].cells[placed.cellIndex];
    const px = cssLengthToPx(td.style.width) ?? cssLengthToPx(td.getAttribute('width'));
    if (px) columnPx[placed.col] = Math.max(columnPx[placed.col] ?? 0, px);
  }
  const columnWidths = columnPx.map(px => Math.round((px ?? DEFAULT_COLUMN_PX) * 15));

  const cellsByRow = rows.map(() => []);
  for (const placed of cells) {
    const td = rows[placed.rowIndex].cells[placed.cellIndex];
    const format = cellFormat(td);
    const style = td.style;
    const content = blocks(td, docx, stats);
    const width = columnWidths.slice(placed.col, placed.col + placed.colspan).reduce((a, b) => a + b, 0);

    if (format.fill) stats.fills++;
    if (placed.rowspan > 1 || placed.colspan > 1) stats.merges++;
    stats.cells++;

    cellsByRow[placed.row].push(new docx.TableCell(defined({
      children: content.length ? content : [new docx.Paragraph({})],
      columnSpan: placed.colspan > 1 ? placed.colspan : undefined,
      rowSpan: placed.rowspan > 1 ? placed.rowspan : undefined,
      width: { size: width, type: docx.WidthType.DXA },
      shading: format.fill ? shading(docx, format.fill) : undefined,
      borders: cellBorders(format, stats),
      verticalAlign: format.vertical === 'middle' ? docx.VerticalAlign.CENTER : format.vertical,
      margins: definedOrNone({
        top: cssLengthToTwips(style.paddingTop),
        bottom: cssLengthToTwips(style.paddingBottom),
        left: cssLengthToTwips(style.paddingLeft),
        right: cssLengthToTwips(style.paddingRight),
      }),
    })));
  }

  const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
  const result = new docx.Table({
    rows: rows.map((tr, r) => {
      const height = cssLengthToTwips(tr.style.height) ?? cssLengthToTwips(tr.getAttribute('height'));
      return new docx.TableRow(defined({
        children: cellsByRow[r],
        height: height ? { value: height, rule: docx.HeightRule.ATLEAST } : undefined,
      }));
    }),
    columnWidths,
    width: { size: totalWidth, type: docx.WidthType.DXA },
    layout: docx.TableLayoutType.FIXED,
    borders: docx.TableBorders.NONE,   // only the borders each cell declares
  });
  tableWidths.set(result, totalWidth);
  return result;
}

/**
 * Builds a Word document from HTML. The `docx` module namespace is passed in.
 * Returns { document, stats }; pack it with docx.Packer.toBlob(document).
 */
export function htmlToDocument(html, docx) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const stats = newStats();
  const children = blocks(doc.body, docx, stats);

  // Wide pasted tables get a landscape page
  const widest = Math.max(0, ...children.map(child => tableWidths.get(child) ?? 0));
  const landscape = widest > USABLE_PORTRAIT_TWIPS;

  const document = new docx.Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },   // Office's defaults
    sections: [{
      properties: landscape ? { page: { size: { orientation: docx.PageOrientation.LANDSCAPE } } } : {},
      children: children.length ? children : [new docx.Paragraph({})],
    }],
  });
  return { document, stats };
}
