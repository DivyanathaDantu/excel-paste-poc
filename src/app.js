import * as docx from 'https://cdn.jsdelivr.net/npm/docx@9.8.1/+esm';
import { inlineOfficeHtml } from './inline-styles.js';
import { htmlToWorkbook } from './html-to-xlsx.js';
import { htmlToDocument } from './html-to-docx.js';

const editor = document.getElementById('editor');
const statusEl = document.getElementById('status');
const statsBody = document.getElementById('stats');
const rawSource = document.getElementById('raw-source');
const inlinedSource = document.getElementById('inlined-source');
const downloadButtons = document.querySelectorAll('[data-download]');

// [stats key, label, which export's count to show]
const ROWS = [
  ['cells', 'Table cells', 'xlsx'],
  ['fills', 'Cell fill colours', 'xlsx'],
  ['fontColors', 'Font colours', 'xlsx'],
  ['bold', 'Bold cells', 'xlsx'],
  ['borders', 'Cells with borders', 'xlsx'],
  ['merges', 'Merged cells', 'xlsx'],
  ['richText', 'Mixed formatting in one cell', 'xlsx'],
  ['numbers', 'Numbers / dates kept as numbers (Excel)', 'xlsx'],
  ['highlights', 'Highlighted text (Word)', 'docx'],
  ['paragraphs', 'Paragraphs (Word text)', 'xlsx'],
];

const EXPORTERS = {
  xlsx: {
    stats: html => htmlToWorkbook(html, ExcelJS).stats,
    blob: async html => new Blob([await htmlToWorkbook(html, ExcelJS).workbook.xlsx.writeBuffer()], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
  },
  docx: {
    stats: html => htmlToDocument(html, docx).stats,
    blob: html => docx.Packer.toBlob(htmlToDocument(html, docx).document),
  },
};

let rawHtml = '';

// Capture phase, so in a real app this runs before the editor's own paste handling
editor.addEventListener('paste', async event => {
  const html = event.clipboardData.getData('text/html');   // read before any await
  if (!html) return;                                       // plain text: normal paste
  event.preventDefault();
  await load(html, 'Pasted from clipboard');
}, true);

async function load(html, label) {
  setStatus('Converting…');
  try {
    rawHtml = html;
    const inlined = await inlineOfficeHtml(html);
    editor.innerHTML = inlined;
    rawSource.textContent = html;
    inlinedSource.textContent = inlined;
    refresh();
    setStatus(`${label}. Edits in the paste area flow into the inlined export.`);
  } catch (error) {
    console.error(error);
    setStatus(`Could not convert this content: ${error.message}`);
  }
}

function refresh() {
  const hasContent = editor.textContent.trim() !== '';
  downloadButtons.forEach(button => { button.disabled = !hasContent; });
  if (!hasContent) {
    statsBody.innerHTML = `<tr><td colspan="3" class="empty">Paste something, or load a sample.</td></tr>`;
    return;
  }
  const stats = {};
  for (const [format, exporter] of Object.entries(EXPORTERS)) {
    stats[format] = {
      raw: exporter.stats(rawHtml || editor.innerHTML),
      inlined: exporter.stats(editor.innerHTML),
    };
  }
  statsBody.innerHTML = ROWS
    .map(([key, label, format]) => [label, stats[format].raw[key], stats[format].inlined[key]])
    .filter(([, raw, inlined]) => raw || inlined)
    .map(([label, raw, inlined]) => {
      const lost = inlined > raw;
      return `<tr>
        <th scope="row">${label}</th>
        <td class="${lost ? 'lost' : ''}">${raw}${lost ? ' <span class="tag">lost</span>' : ''}</td>
        <td class="${lost ? 'kept' : ''}">${inlined}</td>
      </tr>`;
    })
    .join('');
}

async function download(format, variant) {
  const html = variant === 'raw' ? rawHtml || editor.innerHTML : editor.innerHTML;
  try {
    const url = URL.createObjectURL(await EXPORTERS[format].blob(html));
    const link = Object.assign(document.createElement('a'), { href: url, download: `pasted-${variant}.${format}` });
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    console.error(error);
    setStatus(`Could not build the .${format} file: ${error.message}`);
  }
}

function setStatus(text) {
  statusEl.textContent = text;
}

let refreshTimer;
editor.addEventListener('input', () => {
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(() => {
    inlinedSource.textContent = editor.innerHTML;
    refresh();
  }, 300);
});

document.querySelectorAll('[data-sample]').forEach(button => {
  button.addEventListener('click', async () => {
    const name = button.dataset.sample;
    const html = await fetch(`samples/${name}.html`).then(r => r.text());
    await load(html, `Loaded the ${name === 'excel' ? 'Excel' : 'Word'} sample (realistic clipboard HTML)`);
  });
});

document.getElementById('clear').addEventListener('click', () => {
  rawHtml = '';
  editor.innerHTML = '';
  rawSource.textContent = inlinedSource.textContent = '';
  setStatus('');
  refresh();
  editor.focus();
});

downloadButtons.forEach(button => {
  button.addEventListener('click', () => download(button.dataset.download, button.dataset.variant));
});

refresh();
