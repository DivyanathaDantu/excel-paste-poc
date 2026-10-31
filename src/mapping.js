// Pure helpers that turn CSS values into ExcelJS values.
// No DOM access here, so everything in this file is covered by `npm test`.

const NAMED_COLORS = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF',
  yellow: 'FFFF00', orange: 'FFA500', purple: '800080', gray: '808080', grey: '808080',
  silver: 'C0C0C0', maroon: '800000', navy: '000080', teal: '008080', olive: '808000',
  lime: '00FF00', aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF',
  // Office writes these system colours into its clipboard HTML
  windowtext: '000000', window: 'FFFFFF',
};

/** Any CSS colour we expect from Office or getComputedStyle -> 'RRGGBB', or null for none/transparent. */
export function parseColor(css) {
  if (!css) return null;
  const value = css.trim().toLowerCase();
  if (value === 'transparent' || value === 'none' || value === 'auto') return null;
  if (NAMED_COLORS[value]) return NAMED_COLORS[value];

  const hex = value.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/);
  if (hex) {
    const h = hex[1].length === 3 ? [...hex[1]].map(ch => ch + ch).join('') : hex[1];
    return h.toUpperCase();
  }

  const rgb = value.match(/^rgba?\(([^)]+)\)$/);
  if (rgb) {
    const parts = rgb[1].split(/[\s,/]+/).filter(Boolean).map(Number);
    if (parts.length < 3 || parts.slice(0, 3).some(Number.isNaN)) return null;
    if (parts.length >= 4 && parts[3] === 0) return null;
    return parts.slice(0, 3)
      .map(n => Math.round(Math.min(255, Math.max(0, n))).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }
  return null;
}

/** ExcelJS wants colours as ARGB. */
export function toArgb(css) {
  const rgb = parseColor(css);
  return rgb ? `FF${rgb}` : null;
}

const PX_PER_UNIT = { px: 1, pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 };

/** '15.0pt' / '20px' / '.5pt' -> pixels. Relative units (em, %) return null. */
export function cssLengthToPx(css) {
  if (!css) return null;
  const m = String(css).trim().match(/^(-?\d*\.?\d+)(px|pt|in|cm|mm)?$/i);
  if (!m) return null;
  return Number(m[1]) * PX_PER_UNIT[(m[2] || 'px').toLowerCase()];
}

export function cssLengthToPt(css) {
  const px = cssLengthToPx(css);
  return px == null ? null : Math.round(px * 0.75 * 2) / 2;   // nearest half point, like Office
}

const GENERIC_FAMILIES = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

/** 'Calibri, sans-serif' / '"Aptos Narrow"' -> 'Calibri' / 'Aptos Narrow'. */
export function fontName(family) {
  if (!family) return null;
  const first = family.split(',')[0].trim().replace(/^["']|["']$/g, '');
  return first && !GENERIC_FAMILIES.has(first.toLowerCase()) ? first : null;
}

/** One side of a CSS border -> ExcelJS border side, or undefined for no border. */
export function borderFromCss(width, style, color) {
  if (!style || style === 'none' || style === 'hidden') return undefined;
  const px = cssLengthToPx(width) ?? 1;
  if (px <= 0) return undefined;

  let excelStyle;
  if (style === 'double') excelStyle = 'double';
  else if (style === 'dotted') excelStyle = 'dotted';
  else if (style === 'dashed') excelStyle = px >= 1.5 ? 'mediumDashed' : 'dashed';
  else if (px < 1.5) excelStyle = 'thin';      // Excel's .5pt
  else if (px < 2.5) excelStyle = 'medium';    // Excel's 1pt / 1.5pt
  else excelStyle = 'thick';

  return { style: excelStyle, color: { argb: toArgb(color) ?? 'FF000000' } };
}

/** CSS length -> twips (1/20 pt), the unit Word uses for spacing, indents and widths. */
export function cssLengthToTwips(css) {
  const px = cssLengthToPx(css);
  return px == null ? undefined : Math.round(px * 15);
}

/** One side of a CSS border -> docx border { style, size (1/8 pt), color }, or undefined for no border. */
export function docxBorderFromCss(width, style, color) {
  if (!style || style === 'none' || style === 'hidden') return undefined;
  const px = cssLengthToPx(width) ?? 1;
  if (px <= 0) return undefined;
  const eighths = px * 0.75 * 8;
  const docxStyle = ['double', 'dotted', 'dashed'].includes(style) ? style : 'single';
  // CSS double draws two lines in the total width; Word's size is per line.
  // Browsers round Excel's .5pt hairlines down, so single lines never go below Word's usual 1/2pt.
  const min = docxStyle === 'double' ? 2 : 4;
  const size = Math.min(96, Math.max(min, Math.round(docxStyle === 'double' ? eighths / 3 : eighths)));
  return { style: docxStyle, size, color: parseColor(color) ?? '000000' };
}

export function horizontalAlign(textAlign) {
  const value = (textAlign || '').replace(/^-webkit-/, '');
  if (value === 'center') return 'center';
  if (value === 'right' || value === 'end') return 'right';
  if (value === 'justify') return 'justify';
  if (value === 'left') return 'left';
  return undefined;   // start / general: let Excel decide
}

export function verticalAlign(css) {
  if (css === 'top' || css === 'text-top') return 'top';
  if (css === 'middle') return 'middle';
  if (css === 'bottom' || css === 'baseline' || css === 'text-bottom') return 'bottom';
  return undefined;
}

/** Column width in Excel's "characters" unit (Calibri 11 max digit width is 7px). */
export function pxToColumnWidth(px) {
  if (!px || px <= 0) return undefined;
  return Math.max(1, Math.round(((px - 5) / 7) * 100) / 100);
}

/** CSS-escaped string from an mso-* property -> plain text. "\#\,\#\#0\.00" -> "#,##0.00" */
export function cssUnescape(value) {
  return value
    .replace(/\\([0-9a-fA-F]{1,6})\s?/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/\\(.)/g, '$1');
}

const NAMED_NUMBER_FORMATS = {
  'general': null,
  'general number': null,
  'short date': 'm/d/yyyy',
  'medium date': 'd-mmm-yy',
  'long date': 'dddd, mmmm d, yyyy',
  'short time': 'h:mm',
  'medium time': 'h:mm AM/PM',
  'long time': 'h:mm:ss AM/PM',
  'percent': '0.00%',
  'fixed': '0.00',
  'standard': '#,##0.00',
  'scientific': '0.00E+00',
  '@': '@',
};

/** Excel's `mso-number-format` value (as written in the clipboard CSS) -> ExcelJS numFmt, or null for General. */
export function excelNumFmtFromMso(raw) {
  if (!raw) return null;
  let value = raw.trim();
  if (/^".*"$/s.test(value) || /^'.*'$/s.test(value)) value = value.slice(1, -1);
  value = cssUnescape(value).trim();
  const named = NAMED_NUMBER_FORMATS[value.toLowerCase()];
  if (named !== undefined) return named;
  return value || null;
}

const NUMBER_TEXT =
  /^(?<open>\()?(?<sign>-)?(?<cur>[$€£¥])?(?<sign2>-)?(?<int>\d{1,3}(?:,\d{3})+|\d+)(?:\.(?<dec>\d+))?(?<pct>%)?(?<close>\))?$/;

/**
 * Text that looks like a number -> { value, numFmt }, otherwise null (keep it as text).
 * Used when the clipboard did not carry an Excel number format (e.g. Word tables).
 * `xnum` is Excel's x:num attribute: the raw value behind the displayed text.
 */
export function inferNumber(text, xnum) {
  const compact = (text || '').replace(/\s+/g, '');
  const m = compact.match(NUMBER_TEXT);
  if (!m) return null;
  const { open, sign, cur, sign2, int, dec, pct, close } = m.groups;
  if (Boolean(open) !== Boolean(close)) return null;
  const digits = int.replace(/,/g, '');
  if (digits.length > 15) return null;                         // account / phone numbers
  if (digits.length > 1 && digits.startsWith('0') && !int.includes(',')) return null;   // IDs like 007

  let parsed = Number(`${digits}.${dec ?? '0'}`);
  if (open || sign || sign2) parsed = -parsed;
  if (pct) parsed /= 100;

  let value = parsed;
  if (xnum != null && xnum !== '') {
    const raw = Number(xnum);
    if (Number.isNaN(raw)) return null;
    // Displayed text may be rounded; accept the raw value if it rounds to what is shown
    const decimals = (dec?.length ?? 0) + (pct ? 2 : 0);
    if (Math.abs(raw - parsed) > 0.5 * 10 ** -decimals + 1e-9) return null;
    value = raw;
  }

  const decimalsPart = dec ? `.${'0'.repeat(dec.length)}` : '';
  let numFmt = (int.includes(',') ? '#,##0' : '0') + decimalsPart;
  if (pct) numFmt += '%';
  if (cur) numFmt = `"${cur}"${numFmt}`;
  if (open) numFmt = `${numFmt};(${numFmt})`;
  const isPlain = !int.includes(',') && !dec && !pct && !cur && !open;
  return { value, numFmt: isPlain ? null : numFmt };
}

/**
 * Place table cells on a grid, honouring colspan/rowspan.
 * rows: [[{ colspan, rowspan }, ...], ...] -> { cells: [{ rowIndex, cellIndex, row, col, rowspan, colspan }], columnCount }
 * row/col are 0-based grid positions.
 */
export function layoutTable(rows) {
  const taken = [];
  const cells = [];
  let columnCount = 0;
  const isTaken = (r, c) => Boolean(taken[r]?.[c]);

  rows.forEach((row, r) => {
    let c = 0;
    row.forEach((spec, cellIndex) => {
      while (isTaken(r, c)) c++;
      const colspan = Math.max(1, spec.colspan || 1);
      const rowspan = Math.max(1, Math.min(spec.rowspan || 1, rows.length - r));
      for (let dr = 0; dr < rowspan; dr++) {
        taken[r + dr] ??= [];
        for (let dc = 0; dc < colspan; dc++) taken[r + dr][c + dc] = true;
      }
      cells.push({ rowIndex: r, cellIndex, row: r, col: c, rowspan, colspan });
      c += colspan;
      columnCount = Math.max(columnCount, c);
    });
  });
  return { cells, columnCount };
}
