// HTML -> ExcelJS workbook. Formatting is read from inline styles only (see read-styles.js).

import {
  cssLengthToPx, cssLengthToPt, borderFromCss, pxToColumnWidth, excelNumFmtFromMso, inferNumber, layoutTable,
} from './mapping.js';
import { collectRuns, cellFormat, fontFromStyle, newStats, sameFont, walkContent } from './read-styles.js';

/** Neutral font -> ExcelJS font (colours become ARGB). */
const excelFont = font =>
  Object.fromEntries(Object.entries(font).map(([key, value]) => (key === 'color' ? [key, { argb: `FF${value}` }] : [key, value])));

function countFont(font, stats) {
  if (font.color && font.color !== '000000') stats.fontColors++;
  if (font.bold) stats.bold++;
}

const richText = (runs, baseFont) => ({
  richText: runs.map(r => ({ text: r.text, font: excelFont({ ...baseFont, ...r.font }) })),
});

function writeValue(cell, el, stats) {
  const runs = collectRuns(el);
  const text = runs.map(r => r.text).join('');
  const baseFont = fontFromStyle(el.style);
  const msoFormat = el.getAttribute('data-num-format');
  const xnum = el.getAttribute('x:num');
  const numFmt = excelNumFmtFromMso(msoFormat);

  if (numFmt === '@' || el.hasAttribute('x:str')) {
    cell.value = text;
    cell.numFmt = '@';
  } else if (xnum != null && msoFormat) {
    // Excel gave us the raw number and its exact format: use both (works for dates too)
    const value = xnum === '' ? inferNumber(text)?.value : Number(xnum);
    if (value == null || Number.isNaN(value)) cell.value = text;
    else {
      cell.value = value;
      if (numFmt) cell.numFmt = numFmt;
      stats.numbers++;
    }
  } else if (inferNumber(text, xnum)) {
    const { value, numFmt: guessed } = inferNumber(text, xnum);
    cell.value = value;
    if (guessed) cell.numFmt = guessed;
    stats.numbers++;
  } else if (runs.length > 1 && runs.some(r => !sameFont(r.font, runs[0].font))) {
    cell.value = richText(runs, baseFont);
    stats.richText++;
  } else {
    cell.value = text;
  }

  // For single-run content the run's font (e.g. a coloured <span> inside the cell) is the cell font
  const font = runs.length === 1 ? { ...baseFont, ...runs[0].font } : baseFont;
  if (Object.keys(font).length) cell.font = excelFont(font);
  countFont(font, stats);
}

function writeCellFormatting(cell, el, stats) {
  const format = cellFormat(el);
  if (format.fill) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${format.fill}` } };
    stats.fills++;
  }

  const border = {};
  for (const [side, css] of Object.entries(format.borders)) {
    const b = borderFromCss(css.width, css.style, css.color);
    if (b) border[side] = b;
  }
  if (Object.keys(border).length) {
    cell.border = border;
    stats.borders++;
  }

  const alignment = {};
  if (format.horizontal) alignment.horizontal = format.horizontal;
  if (format.vertical) alignment.vertical = format.vertical;
  if (format.wrap) alignment.wrapText = true;
  if (Object.keys(alignment).length) cell.alignment = alignment;
}

function writeTable(ws, table, startRow, stats) {
  const rows = [...table.rows];
  const { cells } = layoutTable(
    rows.map(tr => [...tr.cells].map(td => ({ colspan: td.colSpan, rowspan: td.rowSpan }))),
  );

  rows.forEach((tr, r) => {
    const pt = cssLengthToPt(tr.style.height) ?? cssLengthToPt(tr.getAttribute('height'));
    if (pt) ws.getRow(startRow + r).height = pt;
  });

  for (const placed of cells) {
    const td = rows[placed.rowIndex].cells[placed.cellIndex];
    const r = startRow + placed.row;
    const c = 1 + placed.col;
    const cell = ws.getCell(r, c);
    writeValue(cell, td, stats);
    writeCellFormatting(cell, td, stats);
    stats.cells++;

    if (placed.rowspan > 1 || placed.colspan > 1) {
      ws.mergeCells(r, c, r + placed.rowspan - 1, c + placed.colspan - 1);
      stats.merges++;
    }

    // Column widths from single-column cells; never shrink a column another table widened
    const px = cssLengthToPx(td.style.width) ?? cssLengthToPx(td.getAttribute('width'));
    const width = placed.colspan === 1 ? pxToColumnWidth(px) : undefined;
    const column = ws.getColumn(c);
    if (width && width > (column.width ?? 0)) column.width = width;
  }
  return startRow + rows.length;
}

function writeParagraph(ws, el, row, stats) {
  const runs = collectRuns(el);
  if (!runs.length) return row;
  const cell = ws.getCell(row, 1);
  const prefix = el.tagName === 'LI' ? '• ' : '';
  const baseFont = fontFromStyle(el.style);
  if (runs.length > 1 && runs.some(r => !sameFont(r.font, runs[0].font))) {
    runs[0].text = prefix + runs[0].text;
    cell.value = richText(runs, baseFont);
    stats.richText++;
  } else {
    cell.value = prefix + runs.map(r => r.text).join('');
    const font = { ...baseFont, ...runs[0].font };
    cell.font = excelFont(font);
    countFont(font, stats);
  }
  stats.paragraphs++;
  return row + 1;
}

/**
 * Builds a workbook from HTML. `ExcelJS` is passed in (global from the CDN script in the browser).
 * Returns { workbook, stats } where stats counts what formatting made it into the sheet.
 */
export function htmlToWorkbook(html, ExcelJS) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Pasted');
  const stats = newStats();

  let row = 1;
  walkContent(doc.body, {
    onTable: table => { row = writeTable(ws, table, row, stats) + 1; },   // blank row after each table
    onBlock: el => { row = writeParagraph(ws, el, row, stats); },
  });
  return { workbook, stats };
}
