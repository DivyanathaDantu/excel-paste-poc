import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseColor, toArgb, cssLengthToPx, cssLengthToPt, cssLengthToTwips, fontName, borderFromCss, docxBorderFromCss, horizontalAlign,
  verticalAlign, pxToColumnWidth, cssUnescape, excelNumFmtFromMso, inferNumber, layoutTable,
} from '../src/mapping.js';

test('parseColor handles hex, rgb(), names and Office system colours', () => {
  assert.equal(parseColor('#FFC7CE'), 'FFC7CE');
  assert.equal(parseColor('#abc'), 'AABBCC');
  assert.equal(parseColor('rgb(156, 0, 6)'), '9C0006');
  assert.equal(parseColor('rgba(0, 0, 0, 0)'), null);
  assert.equal(parseColor('rgb(10 20 30 / 0.5)'), '0A141E');
  assert.equal(parseColor('windowtext'), '000000');
  assert.equal(parseColor('transparent'), null);
  assert.equal(parseColor(''), null);
  assert.equal(parseColor('var(--x)'), null);
  assert.equal(toArgb('white'), 'FFFFFFFF');
});

test('CSS lengths convert between px and pt', () => {
  assert.equal(cssLengthToPx('20px'), 20);
  assert.equal(cssLengthToPx('15.0pt'), 20);
  assert.equal(cssLengthToPx('.5pt'), 2 / 3);
  assert.equal(cssLengthToPx('1in'), 96);
  assert.equal(cssLengthToPx('120'), 120);
  assert.equal(cssLengthToPx('1em'), null);
  assert.equal(cssLengthToPt('20px'), 15);
  assert.equal(cssLengthToPt('14.6667px'), 11);
});

test('fontName takes the first real family', () => {
  assert.equal(fontName('Calibri, sans-serif'), 'Calibri');
  assert.equal(fontName('"Aptos Narrow", Arial'), 'Aptos Narrow');
  assert.equal(fontName('sans-serif'), null);
});

test('borderFromCss maps Excel border widths to Excel styles', () => {
  assert.deepEqual(borderFromCss('.5pt', 'solid', 'windowtext'), { style: 'thin', color: { argb: 'FF000000' } });
  assert.equal(borderFromCss('1px', 'solid', '#C00000').style, 'thin');
  assert.equal(borderFromCss('2px', 'solid', 'black').style, 'medium');
  assert.equal(borderFromCss('3px', 'solid', 'black').style, 'thick');
  assert.equal(borderFromCss('2pt', 'double', 'black').style, 'double');
  assert.equal(borderFromCss('1px', 'dashed', 'black').style, 'dashed');
  assert.equal(borderFromCss('1px', 'none', 'black'), undefined);
  assert.equal(borderFromCss('0px', 'solid', 'black'), undefined);
});

test('cssLengthToTwips converts to Word units', () => {
  assert.equal(cssLengthToTwips('1in'), 1440);
  assert.equal(cssLengthToTwips('12pt'), 240);
  assert.equal(cssLengthToTwips('20px'), 300);
  assert.equal(cssLengthToTwips('auto'), undefined);
});

test('docxBorderFromCss maps CSS borders to Word borders', () => {
  assert.deepEqual(docxBorderFromCss('1.33333px', 'solid', 'windowtext'), { style: 'single', size: 8, color: '000000' });
  assert.deepEqual(docxBorderFromCss('.5pt', 'solid', '#C00000'), { style: 'single', size: 4, color: 'C00000' });
  assert.deepEqual(docxBorderFromCss('2pt', 'double', 'black'), { style: 'double', size: 5, color: '000000' });
  assert.equal(docxBorderFromCss('1px', 'dotted', 'black').style, 'dotted');
  assert.equal(docxBorderFromCss('0.5px', 'solid', 'black').size, 4);
  assert.equal(docxBorderFromCss('0.1px', 'double', 'black').size, 2);
  assert.equal(docxBorderFromCss('1px', 'none', 'black'), undefined);
  assert.equal(docxBorderFromCss('0px', 'solid', 'black'), undefined);
});

test('alignment mapping', () => {
  assert.equal(horizontalAlign('center'), 'center');
  assert.equal(horizontalAlign('-webkit-right'), 'right');
  assert.equal(horizontalAlign('start'), undefined);
  assert.equal(verticalAlign('middle'), 'middle');
  assert.equal(verticalAlign('baseline'), 'bottom');
  assert.equal(verticalAlign(''), undefined);
});

test('pxToColumnWidth uses Excel character units', () => {
  assert.equal(pxToColumnWidth(64), 8.43);
  assert.equal(pxToColumnWidth(0), undefined);
});

test('Excel mso-number-format values become ExcelJS numFmt', () => {
  assert.equal(cssUnescape('\\#\\,\\#\\#0\\.00'), '#,##0.00');
  assert.equal(excelNumFmtFromMso('"\\0022$\\0022\\#\\,\\#\\#0\\.00"'), '"$"#,##0.00');
  assert.equal(excelNumFmtFromMso('0%'), '0%');
  assert.equal(excelNumFmtFromMso('"Short Date"'), 'm/d/yyyy');
  assert.equal(excelNumFmtFromMso('"\\@"'), '@');
  assert.equal(excelNumFmtFromMso('General'), null);
  assert.equal(excelNumFmtFromMso('"\\[$$-409\\]\\#\\,\\#\\#0\\.00"'), '[$$-409]#,##0.00');
  assert.equal(excelNumFmtFromMso(null), null);
});

test('inferNumber keeps numbers as numbers with a matching format', () => {
  assert.deepEqual(inferNumber('42'), { value: 42, numFmt: null });
  assert.deepEqual(inferNumber('$1,240,000'), { value: 1240000, numFmt: '"$"#,##0' });
  assert.deepEqual(inferNumber('1,234.50'), { value: 1234.5, numFmt: '#,##0.00' });
  assert.deepEqual(inferNumber('85%'), { value: 0.85, numFmt: '0%' });
  assert.deepEqual(inferNumber('(1,000.00)'), { value: -1000, numFmt: '#,##0.00;(#,##0.00)' });
  assert.deepEqual(inferNumber('-3.5'), { value: -3.5, numFmt: '0.0' });
});

test('inferNumber prefers Excel x:num when it agrees with the displayed text', () => {
  assert.equal(inferNumber('63%', '0.625').value, 0.625);
  assert.equal(inferNumber('$75,999.99', '75999.99').value, 75999.99);
  assert.equal(inferNumber('63%', '0.9'), null);
});

test('inferNumber leaves text alone', () => {
  assert.equal(inferNumber('007'), null);
  assert.equal(inferNumber('1234567890123456'), null);
  assert.equal(inferNumber('12/31/2025'), null);
  assert.equal(inferNumber('On track'), null);
  assert.equal(inferNumber('(12'), null);
  assert.equal(inferNumber(''), null);
});

test('layoutTable places colspan and rowspan cells', () => {
  const { cells, columnCount } = layoutTable([
    [{ colspan: 2 }, { rowspan: 2 }, {}],
    [{}, {}, {}],
  ]);
  assert.equal(columnCount, 4);
  assert.deepEqual(cells.map(c => [c.row, c.col]), [[0, 0], [0, 2], [0, 3], [1, 0], [1, 1], [1, 3]]);
});

test('layoutTable clamps rowspan to the table', () => {
  const { cells } = layoutTable([[{ rowspan: 5 }]]);
  assert.equal(cells[0].rowspan, 1);
});
