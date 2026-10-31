// Reads formatting from HTML for the exporters. Like the converter in the real app, this looks ONLY at
// inline styles (plus <b>/<i>/<u> tags), so raw clipboard HTML shows what gets lost and inlined HTML shows the fix.
// Results are format-neutral (colours as 'RRGGBB', sizes in pt); each exporter maps them to its library.

import { parseColor, cssLengthToPt, fontName, horizontalAlign, verticalAlign } from './mapping.js';

export const BLOCKS = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE']);

export const hasBlockChildren = el =>
  [...el.children].some(child => BLOCKS.has(child.tagName) || child.tagName === 'TABLE');

/** { name, size (pt), color ('RRGGBB'), bold, italic, underline, strike }, only the keys the style sets. */
export function fontFromStyle(style) {
  const font = {};
  const name = fontName(style.fontFamily);
  const size = cssLengthToPt(style.fontSize);
  const color = parseColor(style.color);
  const decoration = `${style.textDecoration} ${style.textDecorationLine}`;
  if (name) font.name = name;
  if (size) font.size = size;
  if (color) font.color = color;
  if (style.fontWeight === 'bold' || Number(style.fontWeight) >= 600) font.bold = true;
  if (style.fontStyle === 'italic') font.italic = true;
  if (decoration.includes('underline')) font.underline = true;
  if (decoration.includes('line-through')) font.strike = true;
  return font;
}

export const sameFont = (a, b) => JSON.stringify(a) === JSON.stringify(b);

const TAG_FONT = {
  B: { bold: true }, STRONG: { bold: true }, I: { italic: true }, EM: { italic: true },
  U: { underline: true }, S: { strike: true }, STRIKE: { strike: true }, DEL: { strike: true },
};

/** Elements between the container and a text node, outermost first. */
function chainTo(node, root) {
  const chain = [];
  for (let el = node.parentElement; el && el !== root; el = el.parentElement) chain.unshift(el);
  return chain;
}

/** Font for a text node: inline styles and <b>/<i>/<u> tags from the container down to the text, inner wins. */
function fontAt(chain) {
  return chain.reduce((font, el) => ({ ...font, ...TAG_FONT[el.tagName], ...fontFromStyle(el.style) }), {});
}

/** Background behind a text node set by an inline element (e.g. a Word highlight), or null. */
function backgroundAt(chain) {
  let background = null;
  for (const el of chain) background = parseColor(el.style.backgroundColor) ?? background;
  return background;
}

/**
 * Text runs inside a cell/paragraph: [{ text, font, background }], where text may contain '\n'
 * and adjacent runs with identical formatting are merged.
 */
export function collectRuns(root) {
  const runs = [];
  const preformatted = root.style.whiteSpace?.startsWith('pre');
  const walk = node => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = preformatted ? node.data : node.data.replace(/\s+/g, ' ');
      if (!text) return;
      const chain = chainTo(node, root);
      const font = fontAt(chain);
      const background = backgroundAt(chain);
      const last = runs.at(-1);
      if (last && sameFont(last.font, font) && last.background === background) last.text += text;
      else runs.push({ text, font, background });
    } else if (node.nodeName === 'BR') {
      if (runs.length) runs.at(-1).text += '\n';
      else runs.push({ text: '\n', font: fontFromStyle(root.style), background: null });
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      node.childNodes.forEach(walk);
    }
  };
  root.childNodes.forEach(walk);

  // Trim like a browser would: around line breaks and at both ends
  runs.forEach(run => { run.text = run.text.replace(/ *\n */g, '\n'); });
  if (runs.length) {
    runs[0].text = runs[0].text.replace(/^\s+/, '');
    runs.at(-1).text = runs.at(-1).text.replace(/\s+$/, '');
  }
  return runs.filter(run => run.text);
}

/**
 * Table cell formatting: { fill, borders: { top: { width, style, color } ... }, horizontal, vertical, wrap }.
 * Border values stay as CSS strings; each exporter maps them to its own border styles.
 */
export function cellFormat(el) {
  const style = el.style;
  const fill = parseColor(style.backgroundColor);
  const borders = {};
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const cap = side[0].toUpperCase() + side.slice(1);
    borders[side] = {
      width: style[`border${cap}Width`],
      style: style[`border${cap}Style`],
      color: style[`border${cap}Color`],
    };
  }
  // Word puts alignment on the paragraph inside the cell rather than on the cell itself
  const innerBlock = el.querySelector(':scope > p, :scope > div');
  return {
    fill: fill === 'FFFFFF' ? null : fill,
    borders,
    horizontal: horizontalAlign(
      style.textAlign || el.getAttribute('align') || innerBlock?.style.textAlign || innerBlock?.getAttribute('align'),
    ),
    vertical: verticalAlign(style.verticalAlign || el.getAttribute('valign')),
    wrap: Boolean(style.whiteSpace) && !['nowrap', 'pre'].includes(style.whiteSpace),
  };
}

/** Visits tables and text blocks in document order: onTable(table), onBlock(el). */
export function walkContent(body, { onTable, onBlock }) {
  const visit = el => {
    if (el.tagName === 'TABLE') onTable(el);
    else if (hasBlockChildren(el)) [...el.children].forEach(visit);
    else if (BLOCKS.has(el.tagName) || el === body) onBlock(el);
  };
  if (hasBlockChildren(body)) [...body.children].forEach(visit);
  else visit(body);
}

export const newStats = () =>
  ({ cells: 0, paragraphs: 0, fills: 0, fontColors: 0, bold: 0, borders: 0, merges: 0, richText: 0, numbers: 0, highlights: 0 });
