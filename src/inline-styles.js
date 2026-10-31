// Paste-time fix: Excel and Word put most formatting in a <style> block (".xl65 { background:#FFC7CE }")
// and only reference it by class. Editors drop that block on paste and converters only read inline
// styles, so the colours vanish. Here we let the browser apply the clipboard's CSS in a hidden iframe,
// then write each element's computed style back onto it as an inline style="...".

import { parseColor, cssLengthToPx } from './mapping.js';

const toHex = css => {
  const rgb = parseColor(css);
  return rgb ? `#${rgb}` : null;
};

/** Width/height the source declared (Excel sets these exactly); rendered sizes depend on installed fonts. */
function declaredPx(el, dimension) {
  return cssLengthToPx(el.style[dimension]) ?? cssLengthToPx(el.getAttribute(dimension));
}

const tablesWithDeclaredWidths = new WeakMap();
function tableDeclaresWidths(table) {
  if (!table) return false;
  if (!tablesWithDeclaredWidths.has(table)) {
    const cells = [...table.querySelectorAll('td, th')];
    tablesWithDeclaredWidths.set(table, cells.some(cell => declaredPx(cell, 'width') != null));
  }
  return tablesWithDeclaredWidths.get(table);
}

function cellWidthPx(el) {
  const declared = declaredPx(el, 'width');
  if (declared != null) return Math.round(declared);
  // Excel only declares widths on the first row; other rows inherit them, so leave those empty
  return tableDeclaresWidths(el.closest('table')) ? null : Math.round(el.getBoundingClientRect().width);
}

function inlineStyleFor(el, cs) {
  const out = [];
  const color = toHex(cs.color);
  const background = toHex(cs.backgroundColor);
  if (color) out.push(`color:${color}`);
  if (background) out.push(`background-color:${background}`);
  out.push(`font-family:${cs.fontFamily}`);
  out.push(`font-size:${Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2}pt`);
  if (Number(cs.fontWeight) >= 600 || cs.fontWeight === 'bold') out.push('font-weight:bold');
  if (cs.fontStyle === 'italic') out.push('font-style:italic');
  if (cs.textDecorationLine && cs.textDecorationLine !== 'none') out.push(`text-decoration:${cs.textDecorationLine}`);

  const align = cs.textAlign.replace(/^-webkit-/, '');
  if (!['start', 'left', 'auto'].includes(align)) out.push(`text-align:${align}`);

  // Spacing does not reach the .xlsx, but keeps the pasted content looking like the source in the editor
  if (/^(P|H[1-6]|LI|UL|OL|BLOCKQUOTE)$/.test(el.tagName)) {
    out.push(`margin:${cs.marginTop} ${cs.marginRight} ${cs.marginBottom} ${cs.marginLeft}`);
    if (parseFloat(cs.textIndent)) out.push(`text-indent:${cs.textIndent}`);   // Word list hanging indents
  }

  if (el.tagName === 'TD' || el.tagName === 'TH') {
    out.push(`padding:${cs.paddingTop} ${cs.paddingRight} ${cs.paddingBottom} ${cs.paddingLeft}`);
    out.push(`vertical-align:${cs.verticalAlign}`);
    out.push(`white-space:${cs.whiteSpace}`);
    const width = cellWidthPx(el);
    if (width != null) out.push(`width:${width}px`);
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const style = cs.getPropertyValue(`border-${side}-style`);
      if (style === 'none' || style === 'hidden') continue;
      const width = cs.getPropertyValue(`border-${side}-width`);
      out.push(`border-${side}:${width} ${style} ${toHex(cs.getPropertyValue(`border-${side}-color`)) ?? '#000000'}`);
    }
  }
  if (el.tagName === 'TR') {
    const height = declaredPx(el, 'height') ?? el.getBoundingClientRect().height;
    out.push(`height:${Math.round(height)}px`);
  }
  return out.join(';');
}

/**
 * Excel keeps number formats in a non-standard `mso-number-format` property. Browsers drop unknown
 * properties when computing styles, so read them from the raw CSS text and keep them in data-num-format.
 */
function collectNumberFormatRules(doc) {
  const rules = [];
  for (const styleEl of doc.querySelectorAll('style')) {
    const css = styleEl.textContent.replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');
    for (const [, selectors, body] of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
      const format = body.match(/mso-number-format\s*:\s*("(?:\\.|[^"])*"|'(?:\\.|[^'])*'|[^;]+)/);
      if (format) rules.push({ selectors: selectors.trim(), format: format[1].trim() });
    }
  }
  return rules;
}

function numberFormatFor(el, rules) {
  const inline = (el.getAttribute('style') || '').match(/mso-number-format\s*:\s*("(?:\\.|[^"])*"|'(?:\\.|[^'])*'|[^;]+)/);
  if (inline) return inline[1].trim();
  let found = null;
  for (const rule of rules) {                       // later rules win, as in the cascade
    try {
      if (el.matches(rule.selectors)) found = rule.format;
    } catch {
      // Office CSS contains selectors browsers do not understand (@page, mso-only); skip them
    }
  }
  return found;
}

/**
 * Clipboard HTML (text/html from Excel or Word) -> HTML whose formatting lives entirely in inline styles.
 * Must run in the paste handler, before the editor sanitises the content and drops the <style> block.
 */
export async function inlineOfficeHtml(html) {
  const iframe = document.createElement('iframe');
  iframe.setAttribute('sandbox', 'allow-same-origin');   // we can read its DOM; its scripts never run
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.cssText = 'position:fixed;left:-100000px;top:0;width:1600px;height:900px;border:0;';
  const loaded = new Promise(resolve => { iframe.onload = resolve; });
  iframe.srcdoc = html;
  document.body.appendChild(iframe);

  try {
    await loaded;
    const doc = iframe.contentDocument;
    const win = iframe.contentWindow;
    const numberRules = collectNumberFormatRules(doc);

    // Compute everything before changing anything, so earlier rewrites cannot affect later reads
    const ignored = 'script, style, meta, link, title';
    const elements = [...doc.body.querySelectorAll('*')].filter(el => !el.matches(ignored));
    const updates = elements.map(el => {
      const cs = win.getComputedStyle(el);
      if (cs.display === 'none') return { el, hidden: true };     // hidden rows/columns
      const isCell = el.tagName === 'TD' || el.tagName === 'TH';
      return {
        el,
        style: inlineStyleFor(el, cs),
        numberFormat: isCell ? numberFormatFor(el, numberRules) : null,
      };
    });

    for (const { el, hidden, style, numberFormat } of updates) {
      if (hidden) { el.remove(); continue; }
      el.setAttribute('style', style);
      el.removeAttribute('class');
      if (numberFormat) el.setAttribute('data-num-format', numberFormat);
    }
    doc.body.querySelectorAll(ignored).forEach(n => n.remove());
    return doc.body.innerHTML.trim();
  } finally {
    iframe.remove();
  }
}
