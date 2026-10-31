// Drop-in paste fix for the DevExtreme HtmlEditor (dxHtmlEditor, tested with 25.1).
// Without it, Excel paste loses all fills, font colours and bold; with it, those survive in the editor's value.
//
//   import { attachOfficePaste } from './devextreme-paste.js';
//   const editor = new DevExpress.ui.dxHtmlEditor(element, { ... });   // or the React/Angular/Vue wrapper's instance
//   const detach = attachOfficePaste(editor);
//
// The editor's table model has no merged cells and keeps only uniform cell borders, so those are still lost.

import { inlineOfficeHtml } from './inline-styles.js';

export function attachOfficePaste(htmlEditor) {
  const quill = htmlEditor.getQuillInstance();

  const onPaste = async event => {
    const html = event.clipboardData?.getData('text/html');   // read before any await
    if (!html) return;                                        // plain text: let the editor handle it
    event.preventDefault();                                   // Quill skips paste events already handled
    const range = quill.getSelection(true);
    const inlined = await inlineOfficeHtml(html);
    if (range.length) quill.deleteText(range.index, range.length, 'user');
    quill.clipboard.dangerouslyPasteHTML(range.index, inlined, 'user');   // 'user' keeps it undoable
  };

  // Capture phase on the editor's container runs before Quill's own paste handler on its content element
  const container = htmlEditor.element();
  const element = container instanceof Element ? container : container[0];   // jQuery integration returns a wrapper
  element.addEventListener('paste', onPaste, true);
  return () => element.removeEventListener('paste', onPaste, true);
}
